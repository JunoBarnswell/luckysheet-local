import { Box, Button, Stack, StatePanel, Text } from "@react-sheets/ui-system";
import { WorkspaceErrorBoundary } from "./components/WorkspaceErrorBoundary";
import { WorkbookBackstageShell } from "./workbooks";
import { WorkbookHubContainer } from "./containers/WorkbookHubContainer";
import { useApplicationServices } from "./ApplicationServicesProvider";
import { useAuthSession, useAuthSnapshot } from "./auth/AuthProvider";
import { navigate, useApplicationRoute } from "./app-routing";
import type { CommandDescriptor } from "@react-sheets/command-runtime";
import { useEffect, useRef, useState } from "react";
import { getInitialSessionPhase, useWorkbookSession, type UiSessionIntent } from "@react-sheets/spreadsheet-app";
import { getInitialLocale, persistLocale, type Locale } from "./i18n";
import { useEditorCommandController } from "./editor/command-controller";
import { EditorShell } from "./editor/EditorShell";

function WorkbookRouteGate({ unitId }: { unitId: string }) {
  const auth = useAuthSession();
  const authSnapshot = useAuthSnapshot();
  const { catalog } = useApplicationServices();
  const [localState, setLocalState] = useState<"checking" | "allowed" | "denied">("checking");
  const shareToken = typeof window !== "undefined" ? new URLSearchParams(window.location.search).get("share")?.trim() : null;

  useEffect(() => {
    if (authSnapshot.phase === "authenticated" || shareToken) {
      setLocalState("allowed");
      return;
    }
    let active = true;
    setLocalState("checking");
    void catalog.open(unitId)
      .then(() => { if (active) setLocalState("allowed"); })
      .catch(() => { if (active) setLocalState("denied"); });
    return () => { active = false; };
  }, [authSnapshot.phase, catalog, shareToken, unitId]);

  if (localState === "checking") return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel kind="loading" title="正在打开工作簿" description="正在确认本地缓存或云端访问权限。" /></Box>;
  if (localState === "denied") {
    const configured = authSnapshot.phase !== "unconfigured";
    return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel actionLabel={configured ? "登录以打开云端文件" : "返回工作簿中心"} kind="error" title={configured ? "需要云端登录" : "云端身份尚未配置"} description={configured ? "未在此浏览器中找到本地工作簿；请登录后访问云端文件。" : "该工作簿不在本地缓存中，且 OIDC 配置不可用。"} onAction={() => configured ? void auth.signIn(`/workbooks/${encodeURIComponent(unitId)}`) : navigate("/workbooks", { replace: true })} /></Box>;
  }
  return <WorkspaceErrorBoundary><EditorRoute key={`${unitId}:${authSnapshot.phase}:${authSnapshot.subject ?? "anonymous"}`} unitId={unitId} onOpenHub={() => navigate("/workbooks")} /></WorkspaceErrorBoundary>;
}

/** Route-level orchestration. Visual responsibilities live in editor/* hosts. */
function EditorRoute({ unitId, onOpenHub }: { unitId: string; onOpenHub: () => void }) {
  const auth = useAuthSession();
  const { catalog, createWorkbookSessionOptions } = useApplicationServices();
  const { session, snapshot: state } = useWorkbookSession({ ...createWorkbookSessionOptions(unitId, auth.getAccessToken), initialPhase: getInitialSessionPhase() });
  const [locale, setLocaleState] = useState<Locale>(() => getInitialLocale());
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [backstageOpen, setBackstageOpen] = useState(false);
  const [backstagePanel, setBackstagePanel] = useState<"info" | "options">("info");
  const previousPanelRef = useRef(state.activePanel);
  const isBusy = state.phase !== "ready";

  const setLocale = (nextLocale: Locale) => { setLocaleState(nextLocale); persistLocale(nextLocale); };
  const dispatchCommand = (descriptor: CommandDescriptor) => session.dispatch(descriptor);
  const dispatchSessionIntent = (intent: UiSessionIntent) => { if (intent.type === "panel.open") setSidebarOpen(true); session.dispatchUiSessionIntent(intent); };
  useEffect(() => { if (previousPanelRef.current !== state.activePanel) setSidebarOpen(true); previousPanelRef.current = state.activePanel; }, [state.activePanel]);

  const controller = useEditorCommandController({ session, state, dispatchCommand, dispatchSessionIntent });
  const copyWorkbookLink = () => { void session.createGuestShareLink("editor"); };
  const saveWorkbook = () => { void session.saveWorkbook("Ribbon save"); };
  const exportXlsx = async () => {
    try {
      await session.saveWorkbook("Export workbook");
      const exported = await catalog.exportXlsx(state.unitId);
      const href = URL.createObjectURL(new Blob([exported.buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const link = document.createElement("a"); link.href = href; link.download = exported.fileName; link.click(); URL.revokeObjectURL(href);
    } catch (cause) { session.notify(cause instanceof Error ? cause.message : "XLSX export failed"); }
  };
  const renameWorkbook = async (name: string) => {
    const normalized = name.trim(); if (!normalized) return;
    try { session.renameWorkbook(normalized); } catch (cause) { session.notify(cause instanceof Error ? cause.message : "Unable to rename workbook"); }
  };
  const importXlsx = () => navigate("/workbooks?dialog=import");

  if (backstageOpen) {
    const syncStatus = state.saveState === "saved" ? "synced" : state.saveState === "saving" || state.saveState === "calculating" ? "syncing" : state.saveState === "conflict" ? "conflict" : state.saveState === "offline" ? "offline" : "error";
    const closeWorkbook = async () => { await session.saveWorkbook("Close workbook"); onOpenHub(); };
    const actions = [
      { id: "info", label: "信息", description: "查看存储、版本与同步信息", icon: "info" as const, onSelect: () => setBackstagePanel("info") },
      { id: "save", label: "保存", description: "提交当前工作簿的保存点", icon: "save" as const, disabled: isBusy, onSelect: () => { void session.saveWorkbook("Backstage save"); } },
      { id: "import", label: "导入", description: "导入为新的工作簿", icon: "upload" as const, disabled: isBusy, onSelect: importXlsx },
      { id: "export", label: "导出", description: "下载当前工作簿的 XLSX 副本", icon: "download" as const, disabled: isBusy, onSelect: () => { void exportXlsx(); } },
      { id: "close", label: "关闭", description: "保存后返回工作簿中心", icon: "x" as const, disabled: isBusy, onSelect: () => { void closeWorkbook(); } },
      { id: "options", label: "选项", description: "语言与工作簿偏好", icon: "settings" as const, onSelect: () => setBackstagePanel("options") },
    ];
    return (
      <WorkbookBackstageShell actions={actions} onBack={() => setBackstageOpen(false)} onHelp={() => session.notify("帮助：导入会创建新的工作簿；云端与本地文件的状态会显示在文件中心。")} onSettings={() => setBackstagePanel("options")} readOnly={!state.permissions.editCell} syncStatus={syncStatus} workbookName={state.workbookName}>
        {backstagePanel === "info" ? (
          <Stack gap="md" className="rounded-xl border border-brand-line bg-white p-6">
            <Text size="lg" weight="semibold">工作簿信息</Text>
            <Stack gap="xs">
              <Text size="sm">文件名：{state.workbookName}</Text><Text size="sm">工作簿 ID：{state.unitId}</Text><Text size="sm">当前权限：{state.shareRole ?? "owner"}</Text><Text size="sm">服务端版本：{state.collabRevision}</Text><Text size="sm">待同步操作：{state.pendingChangeSetCount}</Text><Text size="sm">校验和：{state.persistenceChecksum}</Text><Text size="sm">来源：{state.compatibilityReport ? "导入的 XLSX" : "原生工作簿"}</Text>
            </Stack>
          </Stack>
        ) : (
          <Stack gap="md" className="rounded-xl border border-brand-line bg-white p-6">
            <Text size="lg" weight="semibold">选项</Text><Text size="sm">语言</Text>
            <Box className="flex gap-2"><Button onClick={() => setLocale("zh-CN")} size="sm" variant={locale === "zh-CN" ? "brand" : "outline"}>中文</Button><Button onClick={() => setLocale("en-US")} size="sm" variant={locale === "en-US" ? "brand" : "outline"}>English</Button></Box>
            <Text size="xs" tone="muted">自动保存、自动同步和离线缓存由文件中心的当前用户偏好统一管理。</Text>
          </Stack>
        )}
      </WorkbookBackstageShell>
    );
  }

  return <EditorShell state={state} session={session} locale={locale} isBusy={isBusy} sidebarOpen={sidebarOpen} onSidebarOpenChange={setSidebarOpen} controller={controller} dispatchCommand={dispatchCommand} dispatchSessionIntent={dispatchSessionIntent} setLocale={setLocale} copyWorkbookLink={copyWorkbookLink} saveWorkbook={saveWorkbook} exportXlsx={exportXlsx} importXlsx={importXlsx} renameWorkbook={renameWorkbook} onSetBackstageInfo={() => { setBackstagePanel("info"); setBackstageOpen(true); }} onOpenPrintPreview={() => dispatchSessionIntent({ type: "dialog.open", dialog: "print-preview" })} />;
}

export default function App() {
  const route = useApplicationRoute();
  const auth = useAuthSnapshot();
  useEffect(() => { if (typeof window !== "undefined" && window.location.pathname === "/") navigate("/workbooks", { replace: true }); }, []);
  if (route.kind === "auth-callback" || route.kind === "auth-silent-renew" || auth.phase === "loading") return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel kind="loading" title="正在验证登录状态" description="正在建立云端工作簿会话。" /></Box>;
  if (route.kind === "hub") return <WorkbookHubContainer onOpenWorkbook={(unitId) => navigate(`/workbooks/${encodeURIComponent(unitId)}`)} />;
  if (route.kind === "not-found") return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel actionLabel="返回工作簿中心" kind="error" title="页面不存在" description={`未找到 ${route.pathname}`} onAction={() => navigate("/workbooks", { replace: true })} /></Box>;
  if (route.kind !== "workbook") return <Box as="main" className="flex min-h-screen items-center justify-center bg-white p-8"><StatePanel actionLabel="返回工作簿中心" kind="error" title="路由状态无效" description="无法解析当前工作簿路由。" onAction={() => navigate("/workbooks", { replace: true })} /></Box>;
  return <WorkbookRouteGate unitId={route.unitId} />;
}
