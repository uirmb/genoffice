import { isOfficeProtocolMessage, type OfficeProtocolMessage } from '@genoffice/office-protocol'

export interface OfficeIframeBridgeOptions {
  targetWindow: Window
  targetOrigin: string
  sourceWindow?: Window | undefined
  requestTimeoutMs?: number | undefined
}

export type OfficeMessageHandler<TMessage extends OfficeProtocolMessage> = (
  message: TMessage,
  event: MessageEvent,
) => void

interface PendingRequest {
  expectedType: string
  resolve: (message: OfficeProtocolMessage) => void
  reject: (error: Error) => void
  timer: ReturnType<typeof setTimeout>
}

let requestSequence = 0

export function createOfficeRequestId(prefix = 'office'): string {
  requestSequence += 1
  return `${prefix}-${Date.now().toString(36)}-${requestSequence.toString(36)}`
}

export class OfficeIframeBridge<
  TIncoming extends OfficeProtocolMessage,
  TOutgoing extends OfficeProtocolMessage,
> {
  private readonly targetWindow: Window
  private readonly targetOrigin: string
  private readonly sourceWindow: Window
  private readonly requestTimeoutMs: number
  private readonly handlers = new Set<OfficeMessageHandler<TIncoming>>()
  private readonly pending = new Map<string, PendingRequest>()
  private started = false

  constructor(options: OfficeIframeBridgeOptions) {
    if (!options.targetOrigin || options.targetOrigin === '*') {
      throw new Error('Office iframe bridge requires an explicit targetOrigin')
    }
    this.targetWindow = options.targetWindow
    this.targetOrigin = options.targetOrigin
    this.sourceWindow = options.sourceWindow ?? window
    this.requestTimeoutMs = options.requestTimeoutMs ?? 30_000
  }

  start(): void {
    if (this.started) return
    this.started = true
    this.sourceWindow.addEventListener('message', this.handleMessage)
  }

  destroy(): void {
    if (this.started) {
      this.sourceWindow.removeEventListener('message', this.handleMessage)
      this.started = false
    }
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timer)
      pending.reject(new Error('Office iframe bridge destroyed'))
    }
    this.pending.clear()
    this.handlers.clear()
  }

  subscribe(handler: OfficeMessageHandler<TIncoming>): () => void {
    this.handlers.add(handler)
    return () => this.handlers.delete(handler)
  }

  send(message: TOutgoing, transfer: Transferable[] = []): void {
    this.targetWindow.postMessage(message, this.targetOrigin, transfer)
  }

  request<TResponse extends TIncoming>(
    message: TOutgoing & { requestId: string },
    expectedType: TResponse['type'],
    transfer: Transferable[] = [],
  ): Promise<TResponse> {
    const { requestId } = message
    if (this.pending.has(requestId)) {
      return Promise.reject(new Error(`Duplicate office requestId: ${requestId}`))
    }

    return new Promise<TResponse>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(requestId)
        reject(new Error(`Office request timed out: ${message.type}`))
      }, this.requestTimeoutMs)

      this.pending.set(requestId, {
        expectedType,
        resolve: (response) => resolve(response as TResponse),
        reject,
        timer,
      })

      this.send(message, transfer)
    })
  }

  private readonly handleMessage = (event: MessageEvent): void => {
    if (event.origin !== this.targetOrigin) return
    if (event.source !== this.targetWindow) return
    if (!isOfficeProtocolMessage(event.data)) return

    const message = event.data as TIncoming
    const requestId = 'requestId' in message ? message.requestId : undefined
    if (requestId) {
      const pending = this.pending.get(requestId)
      if (pending && pending.expectedType === message.type) {
        clearTimeout(pending.timer)
        this.pending.delete(requestId)
        pending.resolve(message)
        return
      }
    }

    for (const handler of this.handlers) handler(message, event)
  }
}
