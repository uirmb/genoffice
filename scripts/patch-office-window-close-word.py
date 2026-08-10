from pathlib import Path


def replace_once(path: str, old: str, new: str) -> None:
    p = Path(path)
    text = p.read_text(encoding='utf-8')
    count = text.count(old)
    if count != 1:
        raise RuntimeError(f'{path}: expected 1 match, got {count}: {old[:120]!r}')
    p.write_text(text.replace(old, new, 1), encoding='utf-8')


# 1) Shared Office protocol: Host asks the editor to run its normal exit guard;
# editor either grants the close or explicitly reports that the user cancelled it.
replace_once(
    'packages/office-protocol/src/index.ts',
    "  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:save'\n      requestId: string\n    }\n",
    "  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:save'\n      requestId: string\n    }\n  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:request-close'\n      requestId: string\n      payload: { reason: 'window-close' }\n    }\n",
)
replace_once(
    'packages/office-protocol/src/index.ts',
    "  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:close-request'\n      payload: { reason: 'file-menu' }\n    }\n",
    "  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:close-request'\n      requestId?: string\n      payload: { reason: 'file-menu' | 'window-close' }\n    }\n  | {\n      protocol: typeof OFFICE_PROTOCOL_VERSION\n      type: 'office:close-cancelled'\n      requestId: string\n      payload: { reason: 'user-cancelled' }\n    }\n",
)

# 2) Renderer-facing Web lifecycle hooks. Electron preload does not implement these optional methods.
replace_once(
    'apps/docs/src/shared/ipc.ts',
    "  /** The embedded platform owns the actual plugin/window lifecycle. */\n  requestHostClose?(): Promise<void>\n",
    "  /** The embedded platform owns the actual plugin/window lifecycle. */\n  requestHostClose?(): Promise<void>\n  /** The parent window's × button asked this editor to run the same guarded exit flow. */\n  onHostCloseRequest?(handler: () => void): () => void\n  /** The guarded exit dialog was cancelled; release the parent's pending close state. */\n  cancelHostCloseRequest?(): void\n",
)

# 3) Docs Web adapter: retain correlation until the existing requestExit flow decides.
replace_once(
    'apps/docs/src/web/desktop-api.ts',
    "  let pendingStateRequestId: string | null = null\n  let mode: 'view' | 'edit' = 'edit'\n",
    "  let pendingStateRequestId: string | null = null\n  let pendingHostCloseRequest: { requestId: string; reason: 'window-close' } | null = null\n  let mode: 'view' | 'edit' = 'edit'\n",
)
replace_once(
    'apps/docs/src/web/desktop-api.ts',
    "  const closeSaveHandlers = new Set<VoidHandler>()\n  const aiStreamHandlers = new Set<(chunk: AiStreamChunk) => void>()\n",
    "  const closeSaveHandlers = new Set<VoidHandler>()\n  const hostCloseRequestHandlers = new Set<VoidHandler>()\n  const aiStreamHandlers = new Set<(chunk: AiStreamChunk) => void>()\n",
)
replace_once(
    'apps/docs/src/web/desktop-api.ts',
    "    requestHostClose: async () => {\n      await host.requestClose?.()\n    },\n",
    "    requestHostClose: async () => {\n      if (bridge && pendingHostCloseRequest) {\n        const request = pendingHostCloseRequest\n        pendingHostCloseRequest = null\n        bridge.send({\n          protocol: OFFICE_PROTOCOL_VERSION,\n          type: 'office:close-request',\n          requestId: request.requestId,\n          payload: { reason: request.reason },\n        })\n        return\n      }\n      await host.requestClose?.()\n    },\n    onHostCloseRequest: (handler) => {\n      hostCloseRequestHandlers.add(handler)\n      if (pendingHostCloseRequest) queueMicrotask(handler)\n      return () => hostCloseRequestHandlers.delete(handler)\n    },\n    cancelHostCloseRequest: () => {\n      if (!bridge || !pendingHostCloseRequest) return\n      const request = pendingHostCloseRequest\n      pendingHostCloseRequest = null\n      bridge.send({\n        protocol: OFFICE_PROTOCOL_VERSION,\n        type: 'office:close-cancelled',\n        requestId: request.requestId,\n        payload: { reason: 'user-cancelled' },\n      })\n    },\n",
)
replace_once(
    'apps/docs/src/web/desktop-api.ts',
    "      case 'office:save': {\n",
    "      case 'office:request-close': {\n        if (pendingHostCloseRequest) break\n        pendingHostCloseRequest = {\n          requestId: message.requestId,\n          reason: message.payload.reason,\n        }\n        for (const handler of hostCloseRequestHandlers) handler()\n        break\n      }\n      case 'office:save': {\n",
)
replace_once(
    'apps/docs/src/web/desktop-api.ts',
    "      closeCheckHandlers.clear()\n      closeSaveHandlers.clear()\n      aiStreamHandlers.clear()\n",
    "      closeCheckHandlers.clear()\n      closeSaveHandlers.clear()\n      hostCloseRequestHandlers.clear()\n      pendingHostCloseRequest = null\n      aiStreamHandlers.clear()\n",
)

# 4) React reuses exactly the existing File -> Exit guard.
replace_once(
    'apps/docs/src/renderer/App.tsx',
    "  const discardAndExit = useCallback(() => {\n    setShowExitConfirm(false)\n    void window.desktop.requestHostClose?.()\n  }, [])\n",
    "  useEffect(() => {\n    return window.desktop.onHostCloseRequest?.(requestExit)\n  }, [requestExit])\n\n  const cancelExit = useCallback(() => {\n    setShowExitConfirm(false)\n    window.desktop.cancelHostCloseRequest?.()\n  }, [])\n\n  const discardAndExit = useCallback(() => {\n    setShowExitConfirm(false)\n    void window.desktop.requestHostClose?.()\n  }, [])\n",
)
replace_once(
    'apps/docs/src/renderer/App.tsx',
    "          onCancel={() => setShowExitConfirm(false)}\n",
    "          onCancel={cancelExit}\n",
)

# 5) Adapter regression: preserve requestId on grant and report explicit cancellation.
replace_once(
    'apps/docs/tests/web-desktop-api.test.ts',
    "  it('maps host version conflicts to the existing external-modified save contract', async () => {\n",
    "  it('routes the parent window close request through the guarded editor lifecycle', async () => {\n    const { controller, send, emit, destroy } = createHarness()\n    const requested = vi.fn()\n    const off = controller.desktopApi.onHostCloseRequest?.(requested)\n\n    emit({\n      protocol: OFFICE_PROTOCOL_VERSION,\n      type: 'office:request-close',\n      requestId: 'window-close-1',\n      payload: { reason: 'window-close' },\n    })\n    expect(requested).toHaveBeenCalledTimes(1)\n\n    await controller.desktopApi.requestHostClose?.()\n    expect(send).toHaveBeenCalledWith({\n      protocol: OFFICE_PROTOCOL_VERSION,\n      type: 'office:close-request',\n      requestId: 'window-close-1',\n      payload: { reason: 'window-close' },\n    })\n\n    emit({\n      protocol: OFFICE_PROTOCOL_VERSION,\n      type: 'office:request-close',\n      requestId: 'window-close-2',\n      payload: { reason: 'window-close' },\n    })\n    controller.desktopApi.cancelHostCloseRequest?.()\n    expect(send).toHaveBeenCalledWith({\n      protocol: OFFICE_PROTOCOL_VERSION,\n      type: 'office:close-cancelled',\n      requestId: 'window-close-2',\n      payload: { reason: 'user-cancelled' },\n    })\n\n    off?.()\n    destroy()\n  })\n\n  it('maps host version conflicts to the existing external-modified save contract', async () => {\n",
)

print('Word window-close protocol patch applied')
