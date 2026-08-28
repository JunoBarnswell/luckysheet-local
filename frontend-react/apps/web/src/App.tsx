import { Box, Button, CheckToggle, Inline, Select, Stack, StatePanel, Text, TextInput } from "@react-sheets/ui-system";
import { WorkspaceErrorBoundary } from "./components/WorkspaceErrorBoundary";
import { SaveAsDocumentDialog, WorkbookBackstageShell } from "./workbooks";
import { WorkbookHubContainer } from "./containers/WorkbookHubContainer";
import { useApplicationServices } from "./ApplicationServicesProvider";
import { useAuthSession, useAuthSnapshot } from "./auth/AuthProvider";
import { navigate, useApplicationRoute } from "./app-routing";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import { useEffect, useRef, useState } from "react";
import { getInitialSessionPhase, isWorkbookResolutionError, useWorkbookSession, type UiSessionIntent, type WorkbookResolution } from "@react-sheets/spreadsheet-app";
import { getInitialLocale, persistLocale, type Locale } from "./i18n";
import { useEditorCommandController } from "./editor/command-controller";
import { EditorShell } from "./editor/EditorShell";

function WorkbookRouteGate({ unitId }: { unitId: string }) {
  const auth = useAuthSession();
  const authSnapshot = useAuthSnapshot();
  const { catalog } = useApplicationServices();
  const [localState, setLocalState] = useState<"checking" | "allowed" | "denied">("checking");
  const [resolution, setResolution] = useState<WorkbookResolution | null>(null);
  const [resolutionError, setResolutionError] = useState<Error | null>(null);
  const shareToken = typeof window !== "undefined" ? new URLSearchParams(window.location.search).get("share")?.trim() : null;

  useEffect(() => {
    let active = true;
    const controller = new AbortController();
    setLocalState("checking");
    setResolution(null);
    setResolutionError(null);
    void catalog.resolve(unitId, { signal: controller.signal })
      .then((nextResolution) => { if (active) { setResolution(nextResolution); setLocalState("allowed"); } })
      .catch((error: unknown) => { if (active) { setResolutionError(error instanceof Error ? error : new Error("Workbook resolution failed")); setLocalState("denied"); } });
    return () => { active = false; controller.abort(); };
  }, [authSnapshot.phase, catalog, shareToken, unitId]);

  if (localState === "checking") return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel kind="loading" title="正在打开工作簿" description="正在确认本地缓存或云端访问权限。" /></Box>;
  if (localState === "denied") {
    const canSignIn = authSnapshot.phase !== "authenticated" && authSnapshot.phase !== "unconfigured" && !shareToken;
    const memorySessionReset = isWorkbookResolutionError(resolutionError) && resolutionError.code === "memory-session-reset";
    const title = memorySessionReset ? "内存会话已重置" : canSignIn ? "需要云端登录" : "无法打开工作簿";
    const description = memorySessionReset
      ? "本地工作簿只存在于当前页面的内存会话中；刷新或关闭页面后无法恢复。请返回工作簿中心重新创建或导入。"
      : canSignIn
        ? "当前页面内存会话中没有这个本地工作簿；请登录后打开云端文件。"
        : resolutionError?.message ?? (authSnapshot.phase === "unconfigured" ? "该工作簿不在当前页面内存会话中，且云端服务未配置。" : "工作簿解析失败。");
    return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel actionLabel={memorySessionReset || !canSignIn ? "返回工作簿中心" : "登录以打开云端文件"} kind="error" title={title} description={description} onAction={() => canSignIn && !memorySessionReset ? void auth.signIn(`/workbooks/${encodeURIComponent(unitId)}`) : navigate("/workbooks", { replace: true })} /></Box>;
  }
  if (!resolution) return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel kind="loading" title="正在建立工作簿会话" description="正在交接已解析的工作簿上下文。" /></Box>;
  return <WorkspaceErrorBoundary><EditorRoute key={`${unitId}:${resolution.source}:${resolution.mode}:${resolution.revision}:${resolution.access?.role ?? "local"}`} resolution={resolution} onOpenHub={() => navigate("/workbooks")} /></WorkspaceErrorBoundary>;
}

/** Route-level orchestration. Visual responsibilities live in editor/* hosts. */
function EditorRoute({ resolution, onOpenHub }: { resolution: WorkbookResolution; onOpenHub: () => void }) {
  const unitId = resolution.unitId;
  const auth = useAuthSession();
  const { catalog, createWorkbookSessionOptions } = useApplicationServices();
  const { session, snapshot: state } = useWorkbookSession({
    ...createWorkbookSessionOptions(unitId, auth.getAccessToken, resolution.mode !== "remote"),
    initialPhase: getInitialSessionPhase(),
    resolution,
    onReady: () => catalog.markOpened(resolution),
  });
  const [locale, setLocaleState] = useState<Locale>(() => getInitialLocale());
  const [saveAsOpen, setSaveAsOpen] = useState(false);
  const [saveAsBusy, setSaveAsBusy] = useState(false);
  const initialSelectionApplied = useRef(false);
  const isBusy = state.phase !== "ready" || state.pendingCommandCount > 0;

  useEffect(() => {
    const initialCell = typeof window === "undefined" ? null : new URLSearchParams(window.location.search).get("initialCell");
    if (initialSelectionApplied.current || state.phase !== "ready" || !initialCell) return;
    initialSelectionApplied.current = true;
    session.selectAddress(initialCell);
  }, [session, state.phase]);

  const setLocale = (nextLocale: Locale) => { setLocaleState(nextLocale); persistLocale(nextLocale); };
  const dispatchCommand = (descriptor: CommandDescriptor) => session.dispatch(descriptor);
  const dispatchSessionIntent = (intent: UiSessionIntent) => {
    session.dispatchUiSessionIntent(intent);
  };

  const controller = useEditorCommandController({ session, state, locale, dispatchCommand, dispatchSessionIntent });
  const copyWorkbookLink = () => { void session.createGuestShareLink("editor"); };
  const saveWorkbook = () => { void session.saveWorkbook("Ribbon save"); };
  const exportDocument = async () => {
    try {
      await session.saveWorkbook("Export workbook");
      const exported = await catalog.exportWorkbook(state.unitId);
      const href = URL.createObjectURL(new Blob([exported.buffer], { type: mimeTypeForFileName(exported.fileName) }));
      const link = document.createElement("a"); link.href = href; link.download = exported.fileName; link.click(); URL.revokeObjectURL(href);
    } catch (cause) { session.notify(cause instanceof Error ? cause.message : "文档导出失败"); }
  };
  const renameWorkbook = async (name: string) => {
    const normalized = name.trim(); if (!normalized) return;
    try { session.renameWorkbook(normalized); } catch (cause) { session.notify(cause instanceof Error ? cause.message : "Unable to rename workbook"); }
  };
  const importDocument = () => navigate("/workbooks?dialog=import");
  const saveAsDocument = async (fileName: string) => {
    setSaveAsBusy(true);
    try {
      await session.saveWorkbook("Save before Save As");
      const exported = await catalog.exportWorkbook(state.unitId, { fileName });
      const href = URL.createObjectURL(new Blob([exported.buffer], { type: mimeTypeForFileName(exported.fileName) }));
      const link = document.createElement("a"); link.href = href; link.download = exported.fileName; link.click(); URL.revokeObjectURL(href);
      setSaveAsOpen(false);
    } catch (cause) { session.notify(cause instanceof Error ? cause.message : "另存为失败"); }
    finally { setSaveAsBusy(false); }
  };

  if (state.backstage.open) {
    const syncStatus = state.saveState === "saved" ? "synced" : state.saveState === "saving" || state.saveState === "calculating" ? "syncing" : state.saveState === "conflict" ? "conflict" : state.saveState === "offline" ? "offline" : "error";
    const closeWorkbook = async () => { await session.saveWorkbook("Close workbook"); onOpenHub(); };
    const actions = [
      { id: "info", label: "信息", description: "查看存储、版本与同步信息", icon: "info" as const, onSelect: () => session.setBackstagePanel("info") },
      { id: "save", label: "保存", description: "提交当前工作簿的保存点", icon: "save" as const, disabled: isBusy, onSelect: () => { void session.saveWorkbook("Backstage save"); } },
      { id: "save-as", label: "另存为", description: "选择目标协议导出副本", icon: "save" as const, disabled: isBusy, onSelect: () => setSaveAsOpen(true) },
      { id: "import", label: "打开 / 导入", description: "按原生协议打开为新的工作簿", icon: "upload" as const, disabled: isBusy, onSelect: importDocument },
      { id: "export", label: "导出", description: "下载当前工作簿的原生文档副本", icon: "download" as const, disabled: isBusy, onSelect: () => { void exportDocument(); } },
      { id: "close", label: "关闭", description: "保存后返回工作簿中心", icon: "x" as const, disabled: isBusy, onSelect: () => { void closeWorkbook(); } },
      { id: "options", label: "选项", description: "语言与工作簿偏好", icon: "settings" as const, onSelect: () => session.setBackstagePanel("options") },
    ];
    return (
      <>
      <WorkbookBackstageShell activeActionId={state.backstage.panel === "info" ? "info" : state.backstage.panel === "options" ? "options" : undefined} actions={actions} onBack={() => session.closeBackstage()} onHelp={() => session.notify("帮助：打开 / 导入会创建新的工作簿；另存为只创建目标协议副本；云端与本地文件的状态会显示在文件中心。")} onSettings={() => session.setBackstagePanel("options")} readOnly={!state.permissions.editCell} syncStatus={syncStatus} workbookName={state.workbookName}>
        {state.backstage.panel === "info" ? (
          <Stack gap="md" className="rounded-xl border border-brand-line bg-white p-6">
            <Text size="lg" weight="semibold">工作簿信息</Text>
            <Stack gap="xs">
              <Text size="sm">文件名：{state.workbookName}</Text><Text size="sm">工作簿 ID：{state.unitId}</Text><Text size="sm">当前权限：{state.shareRole ?? "owner"}</Text><Text size="sm">服务端版本：{state.collabRevision}</Text><Text size="sm">待同步操作：{state.pendingChangeSetCount}</Text><Text size="sm">校验和：{state.persistenceChecksum}</Text><Text size="sm">来源：{state.compatibilityReport ? "原生文档导入" : "原生工作簿"}</Text>
            </Stack>
          </Stack>
        ) : (
          <Stack gap="md" className="rounded-xl border border-brand-line bg-white p-6">
            <Text size="lg" weight="semibold">选项</Text><Text size="sm">语言</Text>
            <Box className="flex gap-2"><Button onClick={() => setLocale("zh-CN")} size="sm" variant={locale === "zh-CN" ? "brand" : "outline"}>中文</Button><Button onClick={() => setLocale("en-US")} size="sm" variant={locale === "en-US" ? "brand" : "outline"}>English</Button></Box>
            <Text size="sm" weight="semibold">单元格编辑</Text>
            <CheckToggle label="允许直接在单元格中编辑" checked={state.editingOptions.allowEditDirectly} onChange={(event) => session.setWorkbookEditingOptions({ ...state.editingOptions, allowEditDirectly: event.currentTarget.checked })} />
            <CheckToggle label="按 Enter 后移动选区" checked={state.editingOptions.moveAfterEnter} onChange={(event) => session.setWorkbookEditingOptions({ ...state.editingOptions, moveAfterEnter: event.currentTarget.checked })} />
            <Inline gap="sm">
              <Text size="sm">Enter 方向</Text>
              <Select aria-label="Enter 移动方向" value={state.editingOptions.enterDirection} disabled={!state.editingOptions.moveAfterEnter} onChange={(event) => session.setWorkbookEditingOptions({ ...state.editingOptions, enterDirection: event.currentTarget.value as typeof state.editingOptions.enterDirection })}>
                <option value="down">向下</option><option value="up">向上</option><option value="right">向右</option><option value="left">向左</option>
              </Select>
            </Inline>
            <CheckToggle label="启用公式自动完成" checked={state.editingOptions.formulaAutoComplete} onChange={(event) => session.setWorkbookEditingOptions({ ...state.editingOptions, formulaAutoComplete: event.currentTarget.checked })} />
            <CheckToggle label="启用同列文本自动完成" checked={state.editingOptions.valueAutoComplete} onChange={(event) => session.setWorkbookEditingOptions({ ...state.editingOptions, valueAutoComplete: event.currentTarget.checked })} />
            <Inline gap="sm">
              <Text size="sm">固定小数位</Text>
              <TextInput aria-label="固定小数位" className="w-24" inputMode="numeric" placeholder="关闭" value={state.editingOptions.fixedDecimalPlaces ?? ''} onChange={(event) => {
                const raw = event.currentTarget.value.trim();
                if (raw === '') session.setWorkbookEditingOptions({ ...state.editingOptions, fixedDecimalPlaces: null });
                else if (/^\d+$/.test(raw) && Number(raw) <= 15) session.setWorkbookEditingOptions({ ...state.editingOptions, fixedDecimalPlaces: Number(raw) });
              }} />
            </Inline>
            <Text size="xs" tone="muted">这些选项是工作簿语义，通过命令、历史、协作与持久化统一保存。</Text>
          </Stack>
        )}
      </WorkbookBackstageShell>
      <SaveAsDocumentDialog currentFileName={session.getNativeDocumentFileName() ?? `${state.workbookName}.ssjson`} onClose={() => setSaveAsOpen(false)} onSubmit={(fileName) => { void saveAsDocument(fileName); }} open={saveAsOpen} submitting={saveAsBusy} />
      </>
    );
  }

  return <EditorShell state={state} session={session} locale={locale} isBusy={isBusy} controller={controller} dispatchCommand={dispatchCommand} dispatchSessionIntent={dispatchSessionIntent} setLocale={setLocale} copyWorkbookLink={copyWorkbookLink} saveWorkbook={saveWorkbook} exportDocument={exportDocument} importDocument={importDocument} renameWorkbook={renameWorkbook} onOpenPrintPreview={() => dispatchSessionIntent({ type: "dialog.open", dialog: "print-preview" })} />;
}

function mimeTypeForFileName(fileName: string): string {
  const extension = fileName.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1];
  return extension === 'csv' ? 'text/csv'
    : extension === 'txt' || extension === 'prn' || extension === 'dif' || extension === 'slk' ? 'text/plain'
      : extension === 'xml' ? 'application/xml'
        : extension === 'ods' ? 'application/vnd.oasis.opendocument.spreadsheet'
          : extension === 'sjs' ? 'application/zip'
            : extension === 'ssjson' ? 'application/json'
              : 'application/octet-stream';
}

export default function App() {
  const route = useApplicationRoute();
  const auth = useAuthSnapshot();
  useEffect(() => { if (typeof window !== "undefined" && window.location.pathname === "/") navigate("/workbooks", { replace: true }); }, []);
  if (route.kind === "auth-callback" || route.kind === "auth-silent-renew" || auth.phase === "loading") return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel kind="loading" title="正在验证登录状态" description="正在建立云端工作簿会话。" /></Box>;
  if (route.kind === "hub") return <WorkbookHubContainer onOpenWorkbook={(unitId, options) => {
    const query = options?.initialCell ? `?initialCell=${encodeURIComponent(options.initialCell)}` : "";
    navigate(`/workbooks/${encodeURIComponent(unitId)}${query}`);
  }} />;
  if (route.kind === "not-found") return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel actionLabel="返回工作簿中心" kind="error" title="页面不存在" description={`未找到 ${route.pathname}`} onAction={() => navigate("/workbooks", { replace: true })} /></Box>;
  if (route.kind !== "workbook") return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel actionLabel="返回工作簿中心" kind="error" title="路由状态无效" description="无法解析当前工作簿路由。" onAction={() => navigate("/workbooks", { replace: true })} /></Box>;
  return <WorkbookRouteGate unitId={route.unitId} />;
}
