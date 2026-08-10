from pathlib import Path


def replace_once(path: str, old: str, new: str) -> None:
    p = Path(path)
    text = p.read_text(encoding='utf-8')
    count = text.count(old)
    if count != 1:
        raise RuntimeError(f'{path}: expected 1 match, got {count}: {old[:120]!r}')
    p.write_text(text.replace(old, new, 1), encoding='utf-8')


# Shared protocol (same v1 incremental extension as Word).
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

# Renderer-facing lifecycle hooks (optional so Electron preload remains unchanged).
replace_once(
    'apps/slides/src/shared/ipc.ts',
    "  /** Web Host lifecycle extensions; absent in the legacy Electron preload. */\n  saveHistoryVersion?: () => Promise<{ ok: boolean; error?: string }>\n  exportPptx?: () => Promise<{ ok: boolean; error?: string }>\n  requestHostClose?: () => Promise<void>\n",
    "  /** Web Host lifecycle extensions; absent in the legacy Electron preload. */\n  saveHistoryVersion?: () => Promise<{ ok: boolean; error?: string }>\n  exportPptx?: () => Promise<{ ok: boolean; error?: string }>\n  requestHostClose?: () => Promise<void>\n  /** The parent window's × button asked this editor to run the same guarded exit flow. */\n  onHostCloseRequest?: (handler: () => void) => () => void\n  /** The guarded exit dialog was cancelled; release the parent's pending close state. */\n  cancelHostCloseRequest?: () => void\n",
)

# Slides Web adapter keeps the host request correlated until requestExit decides.
replace_once(
    'apps/slides/src/web/slides-api.ts',
    "  const pendingSaveRequestIds = new Set<string>()\n\n  const languageHandlers",
    "  const pendingSaveRequestIds = new Set<string>()\n  let pendingHostCloseRequest: { requestId: string; reason: 'window-close' } | null = null\n\n  const languageHandlers",
)
replace_once(
    'apps/slides/src/web/slides-api.ts',
    "  const renamedHandlers = new Set<(path: string) => void>()\n\n  const setDirtyState",
    "  const renamedHandlers = new Set<(path: string) => void>()\n  const hostCloseRequestHandlers = new Set<() => void>()\n\n  const setDirtyState",
)
replace_once(
    'apps/slides/src/web/slides-api.ts',
    "    requestHostClose: async () => host.requestClose?.(),\n    isDirty: async () => dirty,\n",
    "    requestHostClose: async () => {\n      if (bridge && pendingHostCloseRequest) {\n        const request = pendingHostCloseRequest\n        pendingHostCloseRequest = null\n        bridge.send({\n          protocol: OFFICE_PROTOCOL_VERSION,\n          type: 'office:close-request',\n          requestId: request.requestId,\n          payload: { reason: request.reason },\n        })\n        return\n      }\n      await host.requestClose?.()\n    },\n    onHostCloseRequest: (handler: () => void) => {\n      hostCloseRequestHandlers.add(handler)\n      if (pendingHostCloseRequest) queueMicrotask(handler)\n      return () => hostCloseRequestHandlers.delete(handler)\n    },\n    cancelHostCloseRequest: () => {\n      if (!bridge || !pendingHostCloseRequest) return\n      const request = pendingHostCloseRequest\n      pendingHostCloseRequest = null\n      bridge.send({\n        protocol: OFFICE_PROTOCOL_VERSION,\n        type: 'office:close-cancelled',\n        requestId: request.requestId,\n        payload: { reason: 'user-cancelled' },\n      })\n    },\n    isDirty: async () => dirty,\n",
)
replace_once(
    'apps/slides/src/web/slides-api.ts',
    "      case 'office:save': {\n",
    "      case 'office:request-close': {\n        if (pendingHostCloseRequest) break\n        pendingHostCloseRequest = {\n          requestId: message.requestId,\n          reason: message.payload.reason,\n        }\n        for (const handler of hostCloseRequestHandlers) handler()\n        break\n      }\n      case 'office:save': {\n",
)
replace_once(
    'apps/slides/src/web/slides-api.ts',
    "      renamedHandlers.clear()\n      pendingSaveRequestIds.clear()\n      session = null\n",
    "      renamedHandlers.clear()\n      hostCloseRequestHandlers.clear()\n      pendingHostCloseRequest = null\n      pendingSaveRequestIds.clear()\n      session = null\n",
)

# React: the external window close reuses requestExit; Cancel explicitly releases the Host.
request_exit = """  const requestExit = useCallback(async () => {\n    const sessionDirty = await window.slidesApi.isDirty()\n    if (!dirty && !sessionDirty) {\n      await window.slidesApi.requestHostClose?.()\n      return\n    }\n    setExitConfirmOpen(true)\n  }, [dirty])\n"""
replace_once(
    'apps/slides/src/renderer/App.tsx',
    request_exit,
    request_exit
    + """\n  useEffect(() => {\n    return window.slidesApi.onHostCloseRequest?.(() => {\n      void requestExit()\n    })\n  }, [requestExit])\n\n  const cancelExit = useCallback(() => {\n    setExitConfirmOpen(false)\n    window.slidesApi.cancelHostCloseRequest?.()\n  }, [])\n""",
)
replace_once(
    'apps/slides/src/renderer/App.tsx',
    '        <div className="modal-backdrop" onClick={() => !exitSaving && setExitConfirmOpen(false)}>\n',
    '        <div className="modal-backdrop" onClick={() => !exitSaving && cancelExit()}>\n',
)
replace_once(
    'apps/slides/src/renderer/App.tsx',
    '              <button disabled={exitSaving} onClick={() => setExitConfirmOpen(false)}>\n',
    '              <button disabled={exitSaving} onClick={cancelExit}>\n',
)

# Adapter regression: grant/cancel preserve correlation; ordinary File -> Exit still uses host.requestClose.
replace_once(
    'apps/slides/tests/web-slides-api.test.ts',
    "  it('delegates history, PPTX export, and close to the Host without replacing the current file identity', async () => {\n",
    "  it('correlates the parent window close request through the guarded Slides lifecycle', async () => {\n    const { controller, requestClose, send, emit } = createHarness()\n    const requested = vi.fn()\n    const off = controller.slidesApi.onHostCloseRequest?.(requested)\n\n    emit({\n      protocol: OFFICE_PROTOCOL_VERSION,\n      type: 'office:request-close',\n      requestId: 'window-close-1',\n      payload: { reason: 'window-close' },\n    })\n    expect(requested).toHaveBeenCalledTimes(1)\n\n    await controller.slidesApi.requestHostClose?.()\n    expect(send).toHaveBeenCalledWith({\n      protocol: OFFICE_PROTOCOL_VERSION,\n      type: 'office:close-request',\n      requestId: 'window-close-1',\n      payload: { reason: 'window-close' },\n    })\n    expect(requestClose).not.toHaveBeenCalled()\n\n    emit({\n      protocol: OFFICE_PROTOCOL_VERSION,\n      type: 'office:request-close',\n      requestId: 'window-close-2',\n      payload: { reason: 'window-close' },\n    })\n    controller.slidesApi.cancelHostCloseRequest?.()\n    expect(send).toHaveBeenCalledWith({\n      protocol: OFFICE_PROTOCOL_VERSION,\n      type: 'office:close-cancelled',\n      requestId: 'window-close-2',\n      payload: { reason: 'user-cancelled' },\n    })\n\n    off?.()\n    controller.destroy()\n  })\n\n  it('delegates history, PPTX export, and close to the Host without replacing the current file identity', async () => {\n",
)

print('PPT window-close protocol patch applied')
