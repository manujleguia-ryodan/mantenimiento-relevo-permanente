import { useAuth } from "@/_core/hooks/useAuth";
import { trpc } from "@/lib/trpc";
import {
  mapExcelRows,
  normalizeExcelHeader,
  rowsFromExcelMatrix,
} from "@shared/excelImport";
import { hasReservation } from "@shared/maintenance";
import {
  AlertTriangle,
  ArrowRight,
  BarChart3,
  BellRing,
  Check,
  ChevronDown,
  ClipboardCheck,
  Clock3,
  FileSpreadsheet,
  Filter,
  HardHat,
  History,
  Import,
  LogOut,
  Menu,
  PackageCheck,
  PanelLeftClose,
  RefreshCw,
  Search,
  Send,
  Settings2,
  ShieldCheck,
  SlidersHorizontal,
  Truck,
  Upload,
  X,
} from "lucide-react";
import {
  type ChangeEvent,
  type DragEvent,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import * as XLSX from "xlsx";
import { toast } from "sonner";

type Guard = "A" | "B" | "C" | "D";
type Status = "SIN_REVISAR" | "RECIBIDA" | "PARCIAL" | "FALTA_RECIBIR";

const GUARDS: Guard[] = ["A", "B", "C", "D"];
const STATUS_META: Record<
  Status,
  { label: string; tone: string; dot: string }
> = {
  SIN_REVISAR: {
    label: "Sin revisar",
    tone: "bg-slate-500/15 text-slate-300 border-slate-500/30",
    dot: "bg-slate-400",
  },
  RECIBIDA: {
    label: "Recibida",
    tone: "bg-emerald-400/10 text-emerald-300 border-emerald-400/30",
    dot: "bg-emerald-400",
  },
  PARCIAL: {
    label: "Parcial",
    tone: "bg-amber-400/10 text-amber-300 border-amber-400/30",
    dot: "bg-amber-400",
  },
  FALTA_RECIBIR: {
    label: "Falta recibir",
    tone: "bg-rose-400/10 text-rose-300 border-rose-400/30",
    dot: "bg-rose-400",
  },
};

function formatDate(value: unknown) {
  if (!value) return "—";
  const date = new Date(value as string | number | Date);
  if (Number.isNaN(date.getTime())) return "—";
  return date.toLocaleString("es-CL", {
    day: "2-digit",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function getImportedFields(sourceData?: string | null): Record<string, string> {
  if (!sourceData) return {};
  try {
    const parsed: unknown = JSON.parse(sourceData);
    return parsed && typeof parsed === "object" && !Array.isArray(parsed)
      ? (parsed as Record<string, string>)
      : {};
  } catch {
    return {};
  }
}

function isPrimaryExcelColumn(key: string) {
  const primaryColumns = new Set([
    "ordendetrabajo",
    "numeroot",
    "nroot",
    "ot",
    "workorder",
    "descripciondeordendetrabajo",
    "descripciondeot",
    "descripcionot",
    "descripcion",
    "shorttext",
    "descripciondeubicaciontecnica",
    "ubicaciontecnica",
    "codigodeequipo",
    "codigoequipo",
    "equipmentcode",
    "equipo",
    "asset",
  ]);
  return primaryColumns.has(normalizeExcelHeader(key));
}

function formatImportedCell(value?: string) {
  if (!value?.trim()) return "—";
  try {
    const parsed: unknown = JSON.parse(value);
    if (Array.isArray(parsed)) {
      return (
        parsed
          .map(item => String(item))
          .filter(Boolean)
          .join(" · ") || "—"
      );
    }
  } catch {
    // Normal cells are plain strings.
  }
  return value;
}

function getIsoWeek(date: Date) {
  const target = new Date(
    Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
  );
  const dayNumber = target.getUTCDay() || 7;
  target.setUTCDate(target.getUTCDate() + 4 - dayNumber);
  const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1));
  return Math.ceil(
    ((target.getTime() - yearStart.getTime()) / 86400000 + 1) / 7
  );
}

function StatusPill({ status }: { status: Status }) {
  const meta = STATUS_META[status];
  return (
    <span
      className={`inline-flex items-center gap-2 rounded-full border px-2.5 py-1 text-[11px] font-bold uppercase tracking-wide ${meta.tone}`}
    >
      <span className={`h-1.5 w-1.5 rounded-full ${meta.dot}`} />
      {meta.label}
    </span>
  );
}

function KpiCard({
  label,
  value,
  detail,
  color,
  icon: Icon,
}: {
  label: string;
  value: string | number;
  detail?: string;
  color: string;
  icon: typeof PackageCheck;
}) {
  return (
    <div className="group relative overflow-hidden rounded-2xl border border-white/[0.07] bg-[#141d1f] p-4 shadow-[0_12px_35px_rgba(0,0,0,0.16)] transition duration-200 hover:-translate-y-0.5 hover:border-white/[0.14]">
      <div
        className={`absolute -right-8 -top-8 h-24 w-24 rounded-full blur-2xl ${color} opacity-20`}
      />
      <div className="relative flex items-start justify-between gap-3">
        <div>
          <div className="text-[10px] font-bold uppercase tracking-[0.18em] text-slate-500">
            {label}
          </div>
          <div className="mt-2 font-display text-3xl font-semibold tracking-tight text-white">
            {value}
          </div>
          {detail && (
            <div className="mt-1 text-xs text-slate-500">{detail}</div>
          )}
        </div>
        <div
          className={`rounded-xl border border-white/[0.07] bg-white/[0.04] p-2.5 ${color.replace("bg-", "text-")}`}
        >
          <Icon className="h-4 w-4" />
        </div>
      </div>
    </div>
  );
}

export default function Home() {
  const { user, loading: authLoading, isAuthenticated, logout } = useAuth();
  const isAdmin = user?.role === "admin";
  const isFacilitator = isAdmin || user?.role === "facilitator";
  const isViewer = user?.role === "viewer";
  const roleLabel = isAdmin
    ? "Administrador"
    : isFacilitator
      ? "Facilitador"
      : isViewer
        ? "Visualizador"
        : "Supervisor";
  const [pin, setPin] = useState("");
  const [pinLoginError, setPinLoginError] = useState<string | null>(null);
  const [selectedWeekId, setSelectedWeekId] = useState<number | null>(null);
  const [guard, setGuard] = useState<Guard>("A");
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState<Status | "TODOS">("TODOS");
  const [warehouseFilter, setWarehouseFilter] = useState("TODOS");
  const [activeSection, setActiveSection] = useState("panel");
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [showRelations, setShowRelations] = useState(false);
  const [showAccessManagement, setShowAccessManagement] = useState(false);
  const [editingRow, setEditingRow] = useState<any | null>(null);
  const [showHistory, setShowHistory] = useState(false);
  const [showHandover, setShowHandover] = useState(false);
  const [handoverTo, setHandoverTo] = useState<Guard>("B");
  const [handoverId, setHandoverId] = useState<number | null>(null);
  const [selectedEquipmentIds, setSelectedEquipmentIds] = useState<number[]>(
    []
  );
  const [relationsDraft, setRelationsDraft] = useState<
    Record<Guard, Guard | "">
  >({ A: "", B: "", C: "", D: "" });
  const [statusDraft, setStatusDraft] = useState<Status>("SIN_REVISAR");
  const [warehouseCodesDraft, setWarehouseCodesDraft] = useState<string[]>([]);
  const [responsibleDraft, setResponsibleDraft] = useState("");
  const [noteDraft, setNoteDraft] = useState("");
  const [importPreview, setImportPreview] = useState<{
    fileName: string;
    rows: Array<{
      otNumber: string;
      equipmentCode: string;
      equipmentDescription?: string;
      description?: string;
      extra?: Record<string, string>;
    }>;
    invalid: number;
    duplicates: number;
  } | null>(null);
  const [importWeekNumber, setImportWeekNumber] = useState(
    String(getIsoWeek(new Date()))
  );
  const [importYear, setImportYear] = useState(
    String(new Date().getFullYear())
  );
  const [isDraggingFile, setIsDraggingFile] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const bootstrap = trpc.maintenance.bootstrap.useQuery(undefined, {
    enabled: isAuthenticated,
  });
  const utils = trpc.useUtils();
  const pinLoginMutation = trpc.auth.pinLogin.useMutation({
    onSuccess: async () => {
      setPin("");
      setPinLoginError(null);
      await utils.auth.me.invalidate();
    },
    onError: error => setPinLoginError(error.message),
  });
  const dashboard = trpc.maintenance.dashboard.useQuery(
    { weekId: selectedWeekId ?? 0, guard },
    { enabled: Boolean(isAuthenticated && selectedWeekId) }
  );
  const history = trpc.maintenance.history.useQuery(
    { workOrderId: editingRow?.workOrder?.id ?? 0 },
    { enabled: Boolean(showHistory && editingRow?.workOrder?.id) }
  );
  const accessUsers = trpc.maintenance.listUsers.useQuery(undefined, {
    enabled: Boolean(isAdmin && showAccessManagement),
  });

  const importMutation = trpc.maintenance.importWeek.useMutation({
    onSuccess: result => {
      toast.success(
        `Importación confirmada: ${result.otCount} OT y ${result.equipmentCount} equipos.`
      );
      setShowImport(false);
      setImportPreview(null);
      void utils.maintenance.bootstrap.invalidate();
      setSelectedWeekId(result.weekId);
    },
    onError: error => toast.error(error.message),
  });
  const assignmentMutation = trpc.maintenance.saveAssignments.useMutation({
    onSuccess: result => {
      setSelectedEquipmentIds(result.selectedEquipmentIds);
      void utils.maintenance.dashboard.invalidate();
      toast.success("Equipos de la guardia actualizados.");
    },
    onError: error => toast.error(error.message),
  });
  const receptionMutation = trpc.maintenance.updateReception.useMutation({
    onSuccess: () => {
      toast.success("Recepción actualizada y auditada.");
      setEditingRow(null);
      void utils.maintenance.dashboard.invalidate();
    },
    onError: error => toast.error(error.message),
  });
  const executionMutation = trpc.maintenance.markOperationalExecution.useMutation({
    onSuccess: result => {
      toast.success(
        result.status === "CERRADA_NOTIFICADA"
          ? "OT cerrada y notificada."
          : "OT reabierta para ejecución."
      );
      void utils.maintenance.dashboard.invalidate();
    },
    onError: error => toast.error(error.message),
  });
  const handoverMutation = trpc.maintenance.generateHandover.useMutation({
    onSuccess: result => {
      setHandoverId(result.handoverId);
      toast.success(`Relevo enviado con ${result.pendingCount} OT pendientes.`);
      void utils.maintenance.latestHandover.invalidate();
    },
    onError: error => toast.error(error.message),
  });
  const confirmHandoverMutation = trpc.maintenance.confirmHandover.useMutation({
    onSuccess: () => {
      toast.success("Relevo confirmado. La trazabilidad quedó registrada.");
      setShowHandover(false);
    },
    onError: error => toast.error(error.message),
  });
  const relationsMutation = trpc.maintenance.saveRelations.useMutation({
    onSuccess: relations => {
      setRelationsDraft(prev => {
        const next = { ...prev };
        for (const relation of relations)
          next[relation.fromGuard as Guard] = relation.toGuard as Guard;
        return next;
      });
      toast.success("Relaciones de relevo guardadas.");
      setShowRelations(false);
      void utils.maintenance.bootstrap.invalidate();
    },
    onError: error => toast.error(error.message),
  });
  const userRoleMutation = trpc.maintenance.updateUserRole.useMutation({
    onSuccess: () => {
      toast.success("Rol de usuario actualizado.");
      void accessUsers.refetch();
    },
    onError: error => toast.error(error.message),
  });

  const weeks = bootstrap.data?.weeks ?? [];
  const allEquipment =
    dashboard.data?.allEquipment ?? bootstrap.data?.equipment ?? [];
  const rows = dashboard.data?.rows ?? [];
  const excelColumns = useMemo(() => {
    const first = rows.find(row => row.workOrder.sourceData)?.workOrder
      .sourceData;
    const available = Object.keys(getImportedFields(first)).filter(
      key => !isPrimaryExcelColumn(key)
    );
    const preferred = [
      ["reserva"],
      ["descripciondeoperacion", "descripcionoperacion"],
    ];
    const selected: string[] = [];
    for (const aliases of preferred) {
      const match = available.find(key => {
        const normalized = normalizeExcelHeader(key);
        return aliases.some(alias => normalized === alias || normalized.includes(alias));
      });
      if (match && !selected.includes(match)) selected.push(match);
    }
    if (selected.length < preferred.length) {
      selected.push(
        ...available
          .filter(key => !selected.includes(key))
          .slice(0, preferred.length - selected.length)
      );
    }
    return selected;
  }, [rows]);
  const importPreviewColumns = importPreview
    ? Object.keys(importPreview.rows[0]?.extra ?? {}).filter(
        key => !isPrimaryExcelColumn(key)
      )
    : [];
  const selectedWeek = weeks.find(week => week.id === selectedWeekId);
  const selectedRelation = relationsDraft[guard];
  const editingRequiresLogistics = Boolean(
    editingRow &&
      (editingRow.requiresLogistics ?? hasReservation(editingRow.workOrder.sourceData))
  );

  useEffect(() => {
    if (!selectedWeekId && weeks[0]) setSelectedWeekId(weeks[0].id);
  }, [selectedWeekId, weeks]);

  useEffect(() => {
    if (dashboard.data) {
      setSelectedEquipmentIds(
        dashboard.data.selectedEquipmentIds.filter(
          id => Number.isSafeInteger(id) && id > 0
        )
      );
      const relation = bootstrap.data?.relations?.find(
        item => item.fromGuard === guard
      );
      if (relation) setHandoverTo(relation.toGuard as Guard);
    }
  }, [dashboard.data, bootstrap.data?.relations, guard]);

  useEffect(() => {
    if (bootstrap.data?.relations) {
      setRelationsDraft(prev => {
        const next = { ...prev };
        for (const relation of bootstrap.data!.relations)
          next[relation.fromGuard as Guard] = relation.toGuard as Guard;
        return next;
      });
    }
  }, [bootstrap.data?.relations]);

  const filteredRows = useMemo(() => {
    const query = search.trim().toLowerCase();
    return rows.filter(row => {
      const state = (row.state?.status ?? "SIN_REVISAR") as Status;
      const warehouse = (row.pendingWarehouses ?? []).map((item: any) => item.warehouseCodeSnapshot).join(",") || "NINGUNO";
      const matchesSearch =
        !query ||
        [
          row.workOrder.otNumber,
          row.equipment?.code,
          row.workOrder.description,
          ...Object.values(getImportedFields(row.workOrder.sourceData)),
        ].some(value =>
          String(value ?? "")
            .toLowerCase()
            .includes(query)
        );
      const matchesStatus = statusFilter === "TODOS" || state === statusFilter;
      const matchesWarehouse =
        warehouseFilter === "TODOS" || warehouse === warehouseFilter;
      return matchesSearch && matchesStatus && matchesWarehouse;
    });
  }, [rows, search, statusFilter, warehouseFilter]);

  const kpis = useMemo(() => {
    const logisticsRows = rows.filter(
      row => row.requiresLogistics ?? hasReservation(row.workOrder.sourceData)
    );
    const total = logisticsRows.length;
    const received = logisticsRows.filter(
      row => row.state?.status === "RECIBIDA"
    ).length;
    const partial = logisticsRows.filter(row => row.state?.status === "PARCIAL").length;
    const missing = logisticsRows.filter(
      row => row.state?.status === "FALTA_RECIBIR"
    ).length;
    const pendingWarehouse1 = logisticsRows.filter(row => (row.pendingWarehouses ?? []).some((item: any) => item.warehouseCodeSnapshot === "CENTRAL")).length;
    const pendingWarehouse2 = logisticsRows.filter(row => (row.pendingWarehouses ?? []).some((item: any) => item.warehouseCodeSnapshot === "FERREYROS")).length;
    const pendingDonaldson = logisticsRows.filter(row => (row.pendingWarehouses ?? []).some((item: any) => item.warehouseCodeSnapshot === "DONALDSON")).length;
    const both = pendingDonaldson;
    return {
      total,
      received,
      partial,
      missing,
      pendingWarehouse1,
      pendingWarehouse2,
      both,
      pendingDonaldson,
      withoutReservation: rows.length - total,
      percent: total ? Math.round((received / total) * 100) : 0,
    };
  }, [rows]);

  function openEdit(row: any) {
    setEditingRow(row);
    setStatusDraft((row.state?.status ?? "SIN_REVISAR") as Status);
    setWarehouseCodesDraft((row.pendingWarehouses ?? []).map((item: any) => item.warehouseCodeSnapshot));
    setResponsibleDraft(row.state?.responsibleName ?? user?.name ?? "");
    setNoteDraft(row.state?.note ?? "");
    setShowHistory(false);
  }

  async function handleFile(file: File) {
    try {
      if (!isFacilitator)
        throw new Error("Solo un Facilitador o Administrador puede importar Excel.");
      const allowedExtensions = [".xlsx", ".xls", ".csv"];
      const fileName = file.name.toLowerCase();
      if (!allowedExtensions.some(extension => fileName.endsWith(extension))) {
        throw new Error(
          "Selecciona un archivo Excel válido (.xlsx, .xls o .csv)."
        );
      }
      if (file.size === 0) throw new Error("El archivo está vacío.");
      if (file.size > 15 * 1024 * 1024)
        throw new Error("El archivo supera el límite de 15 MB.");
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheetName = workbook.SheetNames[0];
      if (!sheetName) throw new Error("El archivo no contiene hojas.");
      const sheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
        header: 1,
        defval: "",
        raw: false,
      }) as unknown[][];
      const rawRows =
        rowsFromExcelMatrix(matrix) ??
        XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
          defval: "",
        });
      const result = mapExcelRows(rawRows);
      if (!result.rows.length)
        throw new Error("No se encontraron filas válidas con OT y equipo.");
      setImportPreview({
        fileName: file.name,
        rows: result.rows,
        invalid: result.invalid,
        duplicates: result.duplicates,
      });
      toast.success(
        `${result.rows.length} OT detectadas. Revisa la vista previa antes de confirmar.`
      );
    } catch (error) {
      toast.error(
        error instanceof Error ? error.message : "No fue posible leer el Excel."
      );
    } finally {
      setIsDraggingFile(false);
    }
  }

  function handleFileInput(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) void handleFile(file);
    event.target.value = "";
  }

  function handleFileDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setIsDraggingFile(false);
    const file = event.dataTransfer.files?.[0];
    if (file) void handleFile(file);
  }

  function submitImport() {
    if (!importPreview) return;
    if (!isFacilitator) {
      toast.error("Solo un Facilitador o Administrador puede importar Excel.");
      return;
    }
    const weekNumber = Number(importWeekNumber);
    const year = Number(importYear);
    if (!weekNumber || weekNumber < 1 || weekNumber > 53 || !year) {
      toast.error("Indica un número de semana y año válidos.");
      return;
    }
    importMutation.mutate({
      label: `SEM ${weekNumber}`,
      weekNumber,
      year,
      guard,
      sourceFileName: importPreview.fileName,
      rows: importPreview.rows,
    });
  }

  function toggleEquipment(id: number) {
    if (!isFacilitator) return;
    if (!Number.isSafeInteger(id) || id <= 0) return;
    const current = selectedEquipmentIds.filter(
      item => Number.isSafeInteger(item) && item > 0
    );
    const next = current.includes(id)
      ? current.filter(item => item !== id)
      : [...current, id];
    setSelectedEquipmentIds(next);
    if (selectedWeekId)
      assignmentMutation.mutate({
        weekId: selectedWeekId,
        guard,
        equipmentIds: next,
      });
  }

  function submitReception() {
    if (!editingRow || !selectedWeekId) return;
    if (!isFacilitator) {
      toast.error("Solo un Facilitador o Administrador puede registrar recepción.");
      return;
    }
    if (!(editingRow.requiresLogistics ?? hasReservation(editingRow.workOrder.sourceData))) {
      toast.error("Esta OT no tiene reserva y no requiere clasificación logística.");
      return;
    }
    receptionMutation.mutate({
      workOrderId: editingRow.workOrder.id,
      weekId: selectedWeekId,
      guard,
      status: statusDraft,
      warehouseCodes: statusDraft === "RECIBIDA" ? [] : warehouseCodesDraft,
      responsibleName: responsibleDraft,
      note: noteDraft,
    });
  }

  function submitHandover() {
    if (!selectedWeekId) return;
    if (!isFacilitator) {
      toast.error("Solo un Facilitador o Administrador puede generar relevo.");
      return;
    }
    handoverMutation.mutate({
      weekId: selectedWeekId,
      fromGuard: guard,
      toGuard: handoverTo,
    });
  }

  function scrollTo(section: string) {
    setActiveSection(section);
    setMobileMenuOpen(false);
    document
      .getElementById(section)
      ?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  if (authLoading) {
    return (
      <div className="min-h-screen bg-[#0a0f10] text-slate-300 grid place-items-center">
        <RefreshCw className="h-5 w-5 animate-spin text-amber-300" />
      </div>
    );
  }

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen overflow-hidden bg-[#0a0f10] text-slate-200">
        <div className="mx-auto flex min-h-screen max-w-6xl items-center px-6 py-12">
          <div className="grid w-full gap-12 lg:grid-cols-[1.05fr_0.95fr] lg:items-center">
            <div>
              <div className="mb-8 flex items-center gap-3">
                <div className="grid h-11 w-11 place-items-center rounded-2xl bg-amber-300 text-[#0a0f10]">
                  <HardHat className="h-6 w-6" />
                </div>
                <span className="font-display text-xl font-semibold tracking-wide">
                  MANTENIMIENTO / RELEVO
                </span>
              </div>
              <div className="mb-5 inline-flex items-center gap-2 rounded-full border border-amber-300/20 bg-amber-300/10 px-3 py-1 text-xs font-semibold uppercase tracking-[0.18em] text-amber-200">
                <span className="h-1.5 w-1.5 rounded-full bg-amber-300" />{" "}
                Control operacional
              </div>
              <h1 className="max-w-2xl font-display text-5xl font-semibold leading-[0.95] tracking-tight text-white sm:text-7xl">
                El relevo crítico,
                <br />
                <span className="text-amber-300">sin perder una OT.</span>
              </h1>
              <p className="mt-7 max-w-xl text-base leading-7 text-slate-400">
                Una capa rápida sobre SAP para importar la semana, revisar la
                recepción global de repuestos y entregar pendientes con
                trazabilidad. Ingresa con el PIN compartido del equipo; no se
                requiere correo.
              </p>
              <form
                className="mt-9 flex max-w-xl flex-col gap-3 sm:flex-row"
                onSubmit={event => {
                  event.preventDefault();
                  setPinLoginError(null);
                  pinLoginMutation.mutate({ pin });
                }}
              >
                <input
                  aria-label="PIN compartido"
                  autoComplete="current-password"
                  className="min-w-0 flex-1 rounded-xl border border-white/10 bg-[#121a1c] px-4 py-3 text-white outline-none transition placeholder:text-slate-500 focus:border-amber-300/60 focus:ring-2 focus:ring-amber-300/20 disabled:opacity-60"
                  maxLength={128}
                  minLength={8}
                  onChange={event => setPin(event.target.value)}
                  placeholder="PIN compartido"
                  required
                  type="password"
                  value={pin}
                />
                <button
                  className="inline-flex items-center justify-center gap-3 rounded-xl bg-amber-300 px-5 py-3 text-sm font-bold text-[#0a0f10] shadow-[0_10px_35px_rgba(245,185,66,0.22)] transition hover:bg-amber-200 active:scale-[0.98] disabled:cursor-not-allowed disabled:opacity-60"
                  disabled={pinLoginMutation.isPending || pin.length < 8}
                  type="submit"
                >
                  {pinLoginMutation.isPending ? (
                    "Validando PIN…"
                  ) : (
                    <>
                      Ingresar con PIN <ArrowRight className="h-4 w-4" />
                    </>
                  )}
                </button>
              </form>
              {pinLoginError && (
                <p className="mt-3 text-sm text-rose-300" role="alert">
                  {pinLoginError}
                </p>
              )}
            </div>
            <div className="relative">
              <div className="absolute -inset-8 rounded-full bg-amber-300/10 blur-3xl" />
              <div className="relative rounded-[2rem] border border-white/10 bg-[#121a1c] p-5 shadow-2xl">
                <div className="flex items-center justify-between border-b border-white/[0.07] pb-4">
                  <div>
                    <div className="text-[10px] uppercase tracking-[0.2em] text-slate-500">
                      Vista operacional
                    </div>
                    <div className="mt-1 font-display text-xl text-white">
                      SEM 38 · GUARDIA A
                    </div>
                  </div>
                  <div className="rounded-xl bg-emerald-400/10 p-2 text-emerald-300">
                    <ShieldCheck className="h-5 w-5" />
                  </div>
                </div>
                <div className="grid grid-cols-3 gap-3 py-5">
                  <div className="rounded-xl bg-white/[0.035] p-3">
                    <div className="text-[10px] uppercase text-slate-500">
                      OT totales
                    </div>
                    <div className="mt-2 font-display text-2xl text-white">
                      35
                    </div>
                  </div>
                  <div className="rounded-xl bg-white/[0.035] p-3">
                    <div className="text-[10px] uppercase text-slate-500">
                      Recibidas
                    </div>
                    <div className="mt-2 font-display text-2xl text-emerald-300">
                      22
                    </div>
                  </div>
                  <div className="rounded-xl bg-white/[0.035] p-3">
                    <div className="text-[10px] uppercase text-slate-500">
                      Pendientes
                    </div>
                    <div className="mt-2 font-display text-2xl text-rose-300">
                      13
                    </div>
                  </div>
                </div>
                <div className="space-y-2.5">
                  <div className="flex items-center justify-between rounded-xl border border-white/[0.06] bg-white/[0.025] px-3 py-3">
                    <span className="font-mono text-xs text-slate-300">
                      OT 452781 · CA26
                    </span>
                    <StatusPill status="RECIBIDA" />
                  </div>
                  <div className="flex items-center justify-between rounded-xl border border-white/[0.06] bg-white/[0.025] px-3 py-3">
                    <span className="font-mono text-xs text-slate-300">
                      OT 452782 · CA26
                    </span>
                    <StatusPill status="PARCIAL" />
                  </div>
                  <div className="flex items-center justify-between rounded-xl border border-white/[0.06] bg-white/[0.025] px-3 py-3">
                    <span className="font-mono text-xs text-slate-300">
                      OT 452783 · CA26
                    </span>
                    <StatusPill status="FALTA_RECIBIR" />
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#0a0f10] text-slate-200">
      <header className="sticky top-0 z-30 border-b border-white/[0.07] bg-[#0a0f10]/90 backdrop-blur-xl lg:hidden">
        <div className="flex items-center justify-between px-4 py-3">
          <button
            className="rounded-lg p-2 text-slate-300 hover:bg-white/5"
            onClick={() => setMobileMenuOpen(true)}
          >
            <Menu className="h-5 w-5" />
          </button>
          <div className="flex items-center gap-2">
            <div className="grid h-8 w-8 place-items-center rounded-lg bg-amber-300 text-[#0a0f10]">
              <HardHat className="h-4 w-4" />
            </div>
            <span className="font-display text-lg text-white">RELEVO</span>
          </div>
          <button
            className="rounded-lg p-2 text-slate-300 hover:bg-white/5"
            onClick={() => logout()}
          >
            <LogOut className="h-4 w-4" />
          </button>
        </div>
      </header>
      <div className="mx-auto flex max-w-[1560px]">
        <aside
          className={`fixed inset-y-0 left-0 z-40 w-[270px] border-r border-white/[0.07] bg-[#0d1415] px-4 py-5 transition-transform lg:sticky lg:top-0 lg:block lg:h-screen lg:translate-x-0 ${mobileMenuOpen ? "translate-x-0" : "-translate-x-full"}`}
        >
          <div className="flex items-center justify-between px-2">
            <div className="flex items-center gap-3">
              <div className="grid h-10 w-10 place-items-center rounded-xl bg-amber-300 text-[#0a0f10]">
                <HardHat className="h-5 w-5" />
              </div>
              <div>
                <div className="font-display text-lg leading-none text-white">
                  MANTENIMIENTO
                </div>
                <div className="mt-1 text-[9px] font-bold uppercase tracking-[0.25em] text-amber-300">
                  Relevo operacional
                </div>
              </div>
            </div>
            <button
              className="rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white lg:hidden"
              onClick={() => setMobileMenuOpen(false)}
            >
              <PanelLeftClose className="h-4 w-4" />
            </button>
          </div>
          <div className="mt-10 px-2 text-[10px] font-bold uppercase tracking-[0.2em] text-slate-600">
            Operación
          </div>
          <nav className="mt-3 space-y-1">
            {[
              { id: "panel", label: "Panel de control", icon: BarChart3 },
              { id: "equipos", label: "Equipos de mi guardia", icon: Truck },
              { id: "ot", label: "OT y recepción", icon: ClipboardCheck },
              { id: "relevo", label: "Generar relevo", icon: Send },
            ].filter(item => item.id !== "relevo" || isFacilitator).map(item => (
              <button
                key={item.id}
                onClick={() => scrollTo(item.id)}
                className={`flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left text-sm transition ${activeSection === item.id ? "bg-amber-300/10 font-semibold text-amber-200" : "text-slate-400 hover:bg-white/[0.04] hover:text-white"}`}
              >
                <item.icon className="h-4 w-4" />
                {item.label}
              </button>
            ))}
          </nav>
          <div className="mt-9 px-2 text-[10px] font-bold uppercase tracking-[0.2em] text-slate-600">
            Administración
          </div>
          <nav className="mt-3 space-y-1">
            {isFacilitator && (
              <button
                onClick={() => setShowImport(true)}
                className="flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left text-sm text-slate-400 transition hover:bg-white/[0.04] hover:text-white"
              >
                <Import className="h-4 w-4" />
                Importar Excel semanal
              </button>
            )}
            {isAdmin && (
              <button
                onClick={() => setShowRelations(true)}
                className="flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left text-sm text-slate-400 transition hover:bg-white/[0.04] hover:text-white"
              >
                <Settings2 className="h-4 w-4" />
                Configurar guardias
              </button>
            )}
            {isAdmin && (
              <button
                onClick={() => setShowAccessManagement(true)}
                className="flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left text-sm text-slate-400 transition hover:bg-white/[0.04] hover:text-white"
              >
                <ShieldCheck className="h-4 w-4" />
                Gestionar accesos
              </button>
            )}
          </nav>
          <div className="absolute bottom-5 left-4 right-4 rounded-2xl border border-white/[0.07] bg-white/[0.025] p-3">
            <div className="flex items-center gap-3">
              <div className="grid h-9 w-9 place-items-center rounded-full bg-slate-700 text-sm font-bold text-white">
                {(user?.name ?? "U").slice(0, 1).toUpperCase()}
              </div>
              <div className="min-w-0 flex-1">
                <div className="truncate text-sm font-semibold text-white">
                  {user?.name ?? "Usuario"}
                </div>
                <div className="text-[10px] uppercase tracking-wider text-slate-500">
                  {roleLabel}
                </div>
              </div>
              <button
                onClick={() => logout()}
                className="rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white"
              >
                <LogOut className="h-4 w-4" />
              </button>
            </div>
          </div>
        </aside>
        {mobileMenuOpen && (
          <button
            aria-label="Cerrar menú"
            className="fixed inset-0 z-30 bg-black/60 lg:hidden"
            onClick={() => setMobileMenuOpen(false)}
          />
        )}
        <main className="min-w-0 flex-1 px-4 pb-12 sm:px-6 lg:px-9 lg:pt-8">
          <div className="mx-auto max-w-[1240px]">
            <section id="panel" className="scroll-mt-24 pt-5 lg:pt-0">
              <div className="flex flex-col gap-5 xl:flex-row xl:items-end xl:justify-between">
                <div>
                  <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.22em] text-amber-300">
                    <span className="h-1.5 w-1.5 rounded-full bg-amber-300" />{" "}
                    Centro de control
                  </div>
                  <h1 className="mt-2 font-display text-4xl font-semibold tracking-tight text-white sm:text-5xl">
                    Relevo de mantenimiento
                  </h1>
                  <p className="mt-2 max-w-2xl text-sm leading-6 text-slate-500">
                    Controla la recepción global de repuestos y entrega
                    únicamente lo que necesita atención.
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  {isFacilitator && (
                    <button
                      onClick={() => setShowImport(true)}
                      className="inline-flex items-center gap-2 rounded-xl border border-amber-300/25 bg-amber-300/10 px-3.5 py-2.5 text-xs font-bold text-amber-200 transition hover:bg-amber-300/15"
                    >
                      <Upload className="h-4 w-4" /> Importar Excel
                    </button>
                  )}
                  {isFacilitator && <button
                    onClick={() => {
                      if (selectedWeekId) setShowHandover(true);
                    }}
                    className="inline-flex items-center gap-2 rounded-xl bg-amber-300 px-3.5 py-2.5 text-xs font-bold text-[#0a0f10] transition hover:bg-amber-200 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    <Send className="h-4 w-4" /> Generar relevo
                  </button>}
                </div>
              </div>
              <div className="mt-8 flex flex-col gap-3 rounded-2xl border border-white/[0.07] bg-[#101718] p-3 sm:flex-row sm:items-center sm:justify-between">
                <div className="flex flex-wrap items-center gap-2">
                  <div className="flex items-center gap-2 rounded-xl bg-white/[0.04] px-3 py-2">
                    <Clock3 className="h-4 w-4 text-amber-300" />
                    <select
                      value={selectedWeekId ?? ""}
                      onChange={event =>
                        setSelectedWeekId(Number(event.target.value))
                      }
                      className="bg-transparent text-sm font-semibold text-white outline-none"
                    >
                      <option value="" disabled>
                        Selecciona semana
                      </option>
                      {weeks.map(week => (
                        <option
                          key={week.id}
                          value={week.id}
                          className="bg-[#11191a]"
                        >
                          {week.label} · {week.year}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="flex items-center gap-1 rounded-xl bg-white/[0.04] p-1">
                    {GUARDS.map(item => (
                      <button
                        key={item}
                        onClick={() => setGuard(item)}
                        className={`rounded-lg px-3 py-1.5 text-xs font-bold transition ${guard === item ? "bg-amber-300 text-[#0a0f10]" : "text-slate-500 hover:text-white"}`}
                      >
                        Guardia {item}
                      </button>
                    ))}
                  </div>
                </div>
                <div className="flex items-center gap-2 text-xs text-slate-500">
                  <span>Guardia actual</span>
                  <span className="rounded-full border border-white/10 px-2 py-1 font-semibold text-slate-300">
                    {guard}
                  </span>
                  {selectedRelation && (
                    <>
                      <ArrowRight className="h-3.5 w-3.5" />
                      <span className="rounded-full border border-emerald-400/20 bg-emerald-400/10 px-2 py-1 font-semibold text-emerald-300">
                        Entrega a {selectedRelation}
                      </span>
                    </>
                  )}
                </div>
              </div>
            </section>

            {!weeks.length ? (
              <div className="mt-6 rounded-2xl border border-dashed border-white/10 bg-[#101718] p-10 text-center">
                <FileSpreadsheet className="mx-auto h-8 w-8 text-amber-300" />
                <h2 className="mt-4 font-display text-2xl text-white">
                  Comienza importando la semana
                </h2>
                <p className="mx-auto mt-2 max-w-md text-sm leading-6 text-slate-500">
                  Sube el Excel semanal, revisa las filas detectadas y luego
                  asigna equipos a la guardia actual.
                </p>
                {isFacilitator && (
                  <button
                    onClick={() => setShowImport(true)}
                    className="mt-5 inline-flex items-center gap-2 rounded-xl bg-amber-300 px-4 py-2.5 text-sm font-bold text-[#0a0f10]"
                  >
                    <Upload className="h-4 w-4" /> Importar primera semana
                  </button>
                )}
              </div>
            ) : (
              <>
                <section className="mt-6 grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
                  <KpiCard
                    label="OT con reserva"
                    value={kpis.total}
                    detail={`${kpis.withoutReservation} sin reserva excluidas`}
                    color="bg-sky-400"
                    icon={ClipboardCheck}
                  />
                  <KpiCard
                    label="Recibidas"
                    value={kpis.received}
                    detail={`${kpis.percent}% completas`}
                    color="bg-emerald-400"
                    icon={PackageCheck}
                  />
                  <KpiCard
                    label="Parciales"
                    value={kpis.partial}
                    detail="Requieren seguimiento"
                    color="bg-amber-400"
                    icon={AlertTriangle}
                  />
                  <KpiCard
                    label="Falta recibir"
                    value={kpis.missing}
                    detail="Atención prioritaria"
                    color="bg-rose-400"
                    icon={BellRing}
                  />
                  <KpiCard
                    label="Recepción total"
                    value={`${kpis.percent}%`}
                    detail="Solo OT completas"
                    color="bg-violet-400"
                    icon={BarChart3}
                  />
                </section>
                <section className="mt-6 grid gap-4 xl:grid-cols-[1.15fr_0.85fr]">
                  <div className="rounded-2xl border border-white/[0.07] bg-[#101718] p-5">
                    <div className="flex items-start justify-between">
                      <div>
                        <div className="text-[10px] font-bold uppercase tracking-[0.18em] text-slate-500">
                          Distribución actual
                        </div>
                        <h2 className="mt-1 font-display text-2xl text-white">
                          Estado de recepción
                        </h2>
                      </div>
                      <div className="rounded-xl bg-white/[0.04] p-2 text-slate-400">
                        <BarChart3 className="h-4 w-4" />
                      </div>
                    </div>
                    <div className="mt-5 flex flex-col items-center gap-6 sm:flex-row">
                      <div
                        className="relative grid h-36 w-36 shrink-0 place-items-center rounded-full"
                        style={{
                          background: `conic-gradient(#34d399 0 ${kpis.total ? (kpis.received / kpis.total) * 100 : 0}%, #fbbf24 0 ${kpis.total ? ((kpis.received + kpis.partial) / kpis.total) * 100 : 0}%, #fb7185 0 100%)`,
                        }}
                      >
                        <div className="grid h-24 w-24 place-items-center rounded-full bg-[#101718]">
                          <div className="text-center">
                            <div className="font-display text-2xl text-white">
                              {kpis.percent}%
                            </div>
                            <div className="text-[9px] uppercase tracking-wider text-slate-500">
                              completa
                            </div>
                          </div>
                        </div>
                      </div>
                      <div className="w-full space-y-3">
                        {(
                          ["RECIBIDA", "PARCIAL", "FALTA_RECIBIR"] as Status[]
                        ).map(status => (
                          <div
                            key={status}
                            className="flex items-center justify-between text-sm"
                          >
                            <div className="flex items-center gap-2">
                              <span
                                className={`h-2 w-2 rounded-full ${STATUS_META[status].dot}`}
                              />
                              <span className="text-slate-400">
                                {STATUS_META[status].label}
                              </span>
                            </div>
                            <span className="font-mono font-semibold text-white">
                              {status === "RECIBIDA"
                                ? kpis.received
                                : status === "PARCIAL"
                                  ? kpis.partial
                                  : kpis.missing}
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                  <div className="rounded-2xl border border-white/[0.07] bg-[#101718] p-5">
                    <div className="text-[10px] font-bold uppercase tracking-[0.18em] text-slate-500">
                      Concentración de pendientes
                    </div>
                    <h2 className="mt-1 font-display text-2xl text-white">
                      Pendientes por almacén
                    </h2>
                    <div className="mt-5 space-y-3">
                      <div className="flex items-center justify-between rounded-xl bg-white/[0.035] p-3">
                        <div className="flex items-center gap-3">
                          <div className="grid h-8 w-8 place-items-center rounded-lg bg-orange-400/10 text-orange-300">
                            <span className="font-mono text-xs">01</span>
                          </div>
                          <span className="text-sm text-slate-300">
                            Central
                          </span>
                        </div>
                        <span className="font-display text-2xl text-white">
                          {kpis.pendingWarehouse1}
                        </span>
                      </div>
                      <div className="flex items-center justify-between rounded-xl bg-white/[0.035] p-3">
                        <div className="flex items-center gap-3">
                          <div className="grid h-8 w-8 place-items-center rounded-lg bg-cyan-400/10 text-cyan-300">
                            <span className="font-mono text-xs">02</span>
                          </div>
                          <span className="text-sm text-slate-300">
                            Ferreyros
                          </span>
                        </div>
                        <span className="font-display text-2xl text-white">
                          {kpis.pendingWarehouse2}
                        </span>
                      </div>
                      <div className="flex items-center justify-between rounded-xl bg-white/[0.035] p-3">
                        <div className="flex items-center gap-3">
                          <div className="grid h-8 w-8 place-items-center rounded-lg bg-violet-400/10 text-violet-300">
                            <span className="font-mono text-xs">AB</span>
                          </div>
                          <span className="text-sm text-slate-300">
                            Donaldson
                          </span>
                        </div>
                        <span className="font-display text-2xl text-white">
                          {kpis.pendingDonaldson}
                        </span>
                      </div>
                    </div>
                  </div>
                </section>

                <section
                  id="equipos"
                  className="mt-6 scroll-mt-24 rounded-2xl border border-white/[0.07] bg-[#101718] p-5"
                >
                  <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
                    <div>
                      <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.18em] text-amber-300">
                        <Truck className="h-3.5 w-3.5" /> Asignación semanal
                      </div>
                      <h2 className="mt-1 font-display text-2xl text-white">
                        Equipos de mi guardia
                      </h2>
                      <p className="mt-1 text-sm text-slate-500">
                        Selecciona los equipos que deben aparecer en el panel de
                        la Guardia {guard}.
                      </p>
                    </div>
                    <div className="rounded-xl border border-amber-300/20 bg-amber-300/10 px-3 py-2 text-xs text-amber-200">
                      {selectedEquipmentIds.length} seleccionados
                    </div>
                  </div>
                  <div className="mt-5 grid grid-cols-2 gap-2 sm:grid-cols-4 lg:grid-cols-6">
                    {allEquipment.map(item => (
                      <button
                        key={item.id}
                        onClick={() => toggleEquipment(item.id)}
                        disabled={!isFacilitator}
                        className={`flex items-center justify-between rounded-xl border px-3 py-3 text-left transition disabled:cursor-default disabled:opacity-70 ${selectedEquipmentIds.includes(item.id) ? "border-amber-300/40 bg-amber-300/10 text-amber-100" : "border-white/[0.07] bg-white/[0.025] text-slate-400 hover:border-white/15 hover:text-white"}`}
                      >
                        <span>
                          <span className="block font-mono text-sm font-semibold">
                            {item.code}
                          </span>
                          <span className="mt-1 block truncate text-[10px] text-slate-500">
                            {item.description || "Equipo operativo"}
                          </span>
                        </span>
                        {selectedEquipmentIds.includes(item.id) && (
                          <span className="grid h-5 w-5 place-items-center rounded-full bg-amber-300 text-[#0a0f10]">
                            <Check className="h-3 w-3" />
                          </span>
                        )}
                      </button>
                    ))}
                    {!allEquipment.length && (
                      <div className="col-span-full rounded-xl border border-dashed border-white/10 p-6 text-center text-sm text-slate-500">
                        Importa una semana para cargar equipos.
                      </div>
                    )}
                  </div>
                </section>

                <section
                  id="ot"
                  className="mt-6 scroll-mt-24 rounded-2xl border border-white/[0.07] bg-[#101718] p-5"
                >
                  <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
                    <div>
                      <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.18em] text-amber-300">
                        <ClipboardCheck className="h-3.5 w-3.5" /> Control
                        operativo
                      </div>
                      <h2 className="mt-1 font-display text-2xl text-white">
                        OT de la guardia {guard}
                      </h2>
                      <p className="mt-1 text-xs text-slate-500">
                        OT, equipo, reserva y descripción de operación.
                      </p>
                    </div>
                    <div className="flex flex-col gap-2 sm:flex-row">
                      <div className="relative">
                        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-600" />
                        <input
                          value={search}
                          onChange={event => setSearch(event.target.value)}
                          placeholder="Buscar OT, equipo o descripción"
                          className="h-10 w-full rounded-xl border border-white/[0.08] bg-white/[0.035] pl-9 pr-3 text-sm text-white outline-none placeholder:text-slate-600 focus:border-amber-300/40 sm:w-64"
                        />
                      </div>
                      <select
                        value={statusFilter}
                        onChange={event =>
                          setStatusFilter(
                            event.target.value as Status | "TODOS"
                          )
                        }
                        className="h-10 rounded-xl border border-white/[0.08] bg-[#151e20] px-3 text-xs text-slate-300 outline-none"
                      >
                        <option value="TODOS">Todos los estados</option>
                        <option value="SIN_REVISAR">Sin revisar</option>
                        <option value="RECIBIDA">Recibida</option>
                        <option value="PARCIAL">Parcial</option>
                        <option value="FALTA_RECIBIR">Falta recibir</option>
                      </select>
                      <select
                        value={warehouseFilter}
                        onChange={event =>
                          setWarehouseFilter(event.target.value)
                        }
                        className="h-10 rounded-xl border border-white/[0.08] bg-[#151e20] px-3 text-xs text-slate-300 outline-none"
                      >
                        <option value="TODOS">Todos los almacenes</option>
                        <option value="CENTRAL">Central</option>
                        <option value="FERREYROS">Ferreyros</option>
                        <option value="DONALDSON">Donaldson</option>
                        <option value="NINGUNO">Sin pendiente</option>
                      </select>
                    </div>
                  </div>
                  <div className="mt-5 overflow-x-auto rounded-xl border border-white/[0.06]">
                    <table className="w-full min-w-max text-left">
                      <thead className="bg-white/[0.035] text-[10px] uppercase tracking-[0.15em] text-slate-600">
                        <tr>
                          <th className="whitespace-nowrap px-4 py-3 font-bold">
                            OT
                          </th>
                          <th className="whitespace-nowrap px-4 py-3 font-bold">
                            Equipo
                          </th>
                          {excelColumns.map(column => (
                            <th
                              key={column}
                              className="whitespace-nowrap px-3 py-3 font-bold"
                            >
                              {normalizeExcelHeader(column).includes("reserva")
                                ? "Reserva"
                                : normalizeExcelHeader(column).includes("operacion")
                                  ? "Descripción de operación"
                                  : column}
                            </th>
                          ))}
                          <th className="whitespace-nowrap px-4 py-3 font-bold">
                            Recepción
                          </th>
                          <th className="whitespace-nowrap px-4 py-3 font-bold">
                            Pendiente
                          </th>
                          <th className="whitespace-nowrap px-4 py-3 font-bold">
                            Responsable
                          </th>
                          <th className="whitespace-nowrap px-4 py-3 font-bold">
                            Actualizado
                          </th>
                          <th className="px-4 py-3" />
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-white/[0.05]">
                        {filteredRows.map(row => {
                          const state = (row.state?.status ??
                            "SIN_REVISAR") as Status;
                          const pending =
                            row.state?.pendingWarehouse1 &&
                            row.state?.pendingWarehouse2
                              ? "Ambos"
                              : row.state?.pendingWarehouse1
                                ? "Almacén 1"
                                : row.state?.pendingWarehouse2
                                  ? "Almacén 2"
                                  : "—";
                          const fields = getImportedFields(
                            row.workOrder.sourceData
                          );
                          const requiresLogistics =
                            row.requiresLogistics ?? hasReservation(row.workOrder.sourceData);
                          return (
                            <tr
                              key={row.workOrder.id}
                              className="group transition hover:bg-white/[0.025]"
                            >
                              <td className="whitespace-nowrap px-4 py-4 font-mono text-sm font-semibold text-white">
                                {row.workOrder.otNumber}
                              </td>
                              <td className="whitespace-nowrap px-4 py-4 text-sm text-amber-200/85">
                                {row.equipment?.code ?? "—"}
                              </td>
                              {excelColumns.map(column => {
                                const value = formatImportedCell(
                                  fields[column]
                                );
                                return (
                                  <td
                                    key={column}
                                    className="max-w-[190px] px-3 py-4 text-xs text-slate-400"
                                  >
                                    <div className="truncate" title={value}>
                                      {value}
                                    </div>
                                  </td>
                                );
                              })}
                              <td className="whitespace-nowrap px-4 py-4">
                                {requiresLogistics ? (
                                  <StatusPill status={state} />
                                ) : (
                                  <span className="text-xs text-slate-500">
                                    Sin reserva
                                  </span>
                                )}
                              </td>
                              <td className="whitespace-nowrap px-4 py-4 text-xs text-slate-400">
                                {requiresLogistics ? (
                                  <>
                                    {pending}
                                    {row.inheritedFrom && <div className="mt-1 text-[10px] text-amber-300">Pendiente heredado de Guardia {row.inheritedFrom}</div>}
                                  </>
                                ) : (
                                  "OT sin reserva (No requiere repuestos)"
                                )}
                              </td>
                              <td className="whitespace-nowrap px-4 py-4 text-xs text-slate-400">
                                {row.state?.responsibleName || "—"}
                              </td>
                              <td className="whitespace-nowrap px-4 py-4 text-xs text-slate-500">
                                {formatDate(row.state?.updatedAt)}
                              </td>
                              <td className="whitespace-nowrap px-4 py-4 text-right">
                                {!isViewer && selectedWeekId && (
                                  <button
                                    onClick={() =>
                                      executionMutation.mutate({
                                        workOrderId: row.workOrder.id,
                                        weekId: selectedWeekId,
                                        guard,
                                        closed: row.execution?.status !== "CERRADA_NOTIFICADA",
                                      })
                                    }
                                    disabled={executionMutation.isPending}
                                    className={`mr-2 rounded-lg border px-2.5 py-1.5 text-[11px] font-semibold transition disabled:opacity-40 ${row.execution?.status === "CERRADA_NOTIFICADA" ? "border-emerald-400/25 text-emerald-300 hover:bg-emerald-400/10" : "border-white/[0.08] text-slate-400 hover:border-emerald-300/30 hover:text-emerald-200"}`}
                                  >
                                    {row.execution?.status === "CERRADA_NOTIFICADA"
                                      ? "Reabrir OT"
                                      : "Cerrar y notificar"}
                                  </button>
                                )}
                                {isFacilitator && requiresLogistics && (
                                  <button
                                    onClick={() => openEdit(row)}
                                    className="rounded-lg border border-white/[0.08] px-2.5 py-1.5 text-[11px] font-semibold text-slate-400 opacity-70 transition hover:border-amber-300/30 hover:text-amber-200 group-hover:opacity-100"
                                  >
                                    Revisar
                                  </button>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                    {!filteredRows.length && (
                      <div className="p-10 text-center">
                        <Filter className="mx-auto h-6 w-6 text-slate-600" />
                        <div className="mt-3 text-sm font-semibold text-slate-400">
                          No hay OT para este filtro
                        </div>
                        <div className="mt-1 text-xs text-slate-600">
                          Selecciona equipos o cambia los filtros para
                          continuar.
                        </div>
                      </div>
                    )}
                  </div>
                  <div className="mt-3 flex items-center justify-between text-[11px] text-slate-600">
                    <span>
                      Mostrando {filteredRows.length} de {rows.length} OT
                      visibles
                    </span>
                    <span>
                      Click en “Revisar” para registrar recepción e historial
                    </span>
                  </div>
                </section>

                {isFacilitator && <section
                  id="relevo"
                  className="mt-6 scroll-mt-24 rounded-2xl border border-amber-300/15 bg-gradient-to-br from-amber-300/[0.08] to-transparent p-5"
                >
                  <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.18em] text-amber-300">
                        <Send className="h-3.5 w-3.5" /> Cierre de guardia
                      </div>
                      <h2 className="mt-1 font-display text-2xl text-white">
                        Listo para generar el relevo
                      </h2>
                      <p className="mt-1 text-sm text-slate-400">
                        Se incluirán solo las OT parciales y faltantes de la
                        vista actual.
                      </p>
                    </div>
                    {isFacilitator && (
                    <button
                      onClick={() => setShowHandover(true)}
                      className="inline-flex items-center justify-center gap-2 rounded-xl bg-amber-300 px-4 py-2.5 text-sm font-bold text-[#0a0f10] transition hover:bg-amber-200"
                    >
                      <Send className="h-4 w-4" /> Generar relevo
                    </button>
                    )}
                  </div>
                  <div className="mt-5 grid gap-2 sm:grid-cols-3">
                    <div className="rounded-xl border border-white/[0.07] bg-black/10 p-3">
                      <div className="text-[10px] uppercase tracking-widest text-slate-500">
                        Entrega
                      </div>
                      <div className="mt-1 font-display text-xl text-white">
                        Guardia {guard}
                      </div>
                    </div>
                    <div className="rounded-xl border border-white/[0.07] bg-black/10 p-3">
                      <div className="text-[10px] uppercase tracking-widest text-slate-500">
                        Recibe
                      </div>
                      <div className="mt-1 font-display text-xl text-emerald-300">
                        Guardia {selectedRelation || "—"}
                      </div>
                    </div>
                    <div className="rounded-xl border border-white/[0.07] bg-black/10 p-3">
                      <div className="text-[10px] uppercase tracking-widest text-slate-500">
                        OT atención
                      </div>
                      <div className="mt-1 font-display text-xl text-rose-300">
                        {kpis.partial + kpis.missing}
                      </div>
                    </div>
                  </div>
                </section>}
              </>
            )}
          </div>
        </main>
      </div>

      {showImport && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-black/75 p-4 backdrop-blur-sm">
          <div className="max-h-[90vh] w-full max-w-3xl overflow-y-auto rounded-2xl border border-white/10 bg-[#11191a] shadow-2xl">
            <div className="flex items-start justify-between border-b border-white/[0.07] p-5">
              <div>
                <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.18em] text-amber-300">
                  <FileSpreadsheet className="h-3.5 w-3.5" /> Importación
                  protegida
                </div>
                <h2 className="mt-1 font-display text-2xl text-white">
                  Importar Excel semanal
                </h2>
                <p className="mt-1 text-sm text-slate-500">
                  Se conserva el archivo original y se valida antes de guardar.
                </p>
              </div>
              <button
                onClick={() => {
                  setShowImport(false);
                  setImportPreview(null);
                }}
                className="rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white"
              >
                <X className="h-5 w-5" />
              </button>
            </div>
            <div className="space-y-5 p-5">
              <div className="grid gap-3 sm:grid-cols-[1fr_1fr_1.4fr]">
                <label className="text-xs text-slate-400">
                  Número de semana
                  <input
                    value={importWeekNumber}
                    onChange={event => setImportWeekNumber(event.target.value)}
                    type="number"
                    min="1"
                    max="53"
                    className="mt-2 h-10 w-full rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 text-sm text-white outline-none focus:border-amber-300/40"
                  />
                </label>
                <label className="text-xs text-slate-400">
                  Año
                  <input
                    value={importYear}
                    onChange={event => setImportYear(event.target.value)}
                    type="number"
                    className="mt-2 h-10 w-full rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 text-sm text-white outline-none focus:border-amber-300/40"
                  />
                </label>
                <div className="flex items-end">
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    className="hidden"
                    onChange={handleFileInput}
                  />
                  <button
                    onClick={() => fileInputRef.current?.click()}
                    className="flex h-10 w-full items-center justify-center gap-2 rounded-xl border border-dashed border-amber-300/30 bg-amber-300/[0.06] text-xs font-bold text-amber-200 transition hover:bg-amber-300/10"
                  >
                    <Upload className="h-4 w-4" />{" "}
                    {importPreview ? "Cambiar archivo" : "Seleccionar Excel"}
                  </button>
                </div>
              </div>
              <div
                onDragEnter={event => {
                  event.preventDefault();
                  setIsDraggingFile(true);
                }}
                onDragOver={event => event.preventDefault()}
                onDragLeave={event => {
                  event.preventDefault();
                  setIsDraggingFile(false);
                }}
                onDrop={handleFileDrop}
                className={`rounded-2xl border border-dashed p-5 text-center transition ${isDraggingFile ? "border-amber-300 bg-amber-300/10" : "border-white/10 bg-white/[0.02] hover:border-amber-300/30"}`}
              >
                <div className="flex items-center justify-center gap-2 text-sm font-semibold text-slate-300">
                  <Upload className="h-4 w-4 text-amber-300" /> Arrastra aquí tu
                  Excel semanal
                </div>
                <div className="mt-1 text-xs text-slate-600">
                  También puedes usar el botón de selección · máximo 15 MB ·
                  .xlsx, .xls o .csv
                </div>
                {isDraggingFile && (
                  <div className="mt-2 text-xs font-bold text-amber-200">
                    Suelta el archivo para comenzar la lectura
                  </div>
                )}
              </div>
              {importPreview ? (
                <div className="rounded-2xl border border-white/[0.07] bg-black/10">
                  <div className="flex flex-col gap-2 border-b border-white/[0.07] p-4 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <div className="font-mono text-sm text-white">
                        {importPreview.fileName}
                      </div>
                      <div className="mt-1 text-xs text-slate-500">
                        {importPreview.rows.length} OT detectadas ·{" "}
                        {importPreview.invalid} filas inválidas ·{" "}
                        {importPreview.duplicates} repeticiones de OT detectadas
                      </div>
                    </div>
                    <div className="rounded-lg bg-emerald-400/10 px-2.5 py-1.5 text-[11px] font-bold text-emerald-300">
                      Vista previa lista
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full min-w-max text-left text-xs">
                      <thead className="bg-white/[0.03] uppercase tracking-wider text-slate-600">
                        <tr>
                          <th className="px-4 py-2.5">OT</th>
                          <th className="px-4 py-2.5">Equipo</th>
                          <th className="px-4 py-2.5">Descripción</th>
                          {importPreviewColumns.map(column => (
                            <th
                              key={column}
                              className="whitespace-nowrap px-3 py-2.5"
                            >
                              {column}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-white/[0.05]">
                        {importPreview.rows.slice(0, 7).map((row, index) => (
                          <tr key={`${row.otNumber}-${index}`}>
                            <td className="whitespace-nowrap px-4 py-3 font-mono text-white">
                              {row.otNumber}
                            </td>
                            <td className="whitespace-nowrap px-4 py-3 text-amber-200">
                              {row.equipmentCode}
                            </td>
                            <td className="max-w-[310px] truncate px-4 py-3 text-slate-400">
                              {row.description || "—"}
                            </td>
                            {importPreviewColumns.map(column => {
                              const value = formatImportedCell(
                                row.extra?.[column]
                              );
                              return (
                                <td
                                  key={column}
                                  className="max-w-[190px] px-3 py-3 text-slate-400"
                                >
                                  <div className="truncate" title={value}>
                                    {value}
                                  </div>
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {importPreview.rows.length > 7 && (
                    <div className="border-t border-white/[0.06] px-4 py-3 text-[11px] text-slate-600">
                      Mostrando 7 filas de la vista previa. Se importarán las{" "}
                      {importPreview.rows.length} válidas.
                    </div>
                  )}
                </div>
              ) : (
                <div className="rounded-2xl border border-dashed border-white/10 p-8 text-center">
                  <FileSpreadsheet className="mx-auto h-7 w-7 text-slate-600" />
                  <div className="mt-3 text-sm font-semibold text-slate-400">
                    Aún no hay un archivo seleccionado
                  </div>
                  <div className="mt-1 text-xs text-slate-600">
                    Se reconocen las cabeceras del programa semanal y sus
                    columnas de días.
                  </div>
                </div>
              )}
              <div className="flex flex-col-reverse gap-2 sm:flex-row sm:justify-end">
                <button
                  onClick={() => {
                    setShowImport(false);
                    setImportPreview(null);
                  }}
                  className="rounded-xl px-4 py-2.5 text-sm text-slate-400 hover:bg-white/5 hover:text-white"
                >
                  Cancelar
                </button>
                <button
                  disabled={
                    !importPreview || importMutation.isPending || !isFacilitator
                  }
                  onClick={submitImport}
                  className="inline-flex items-center justify-center gap-2 rounded-xl bg-amber-300 px-4 py-2.5 text-sm font-bold text-[#0a0f10] disabled:cursor-not-allowed disabled:opacity-40"
                >
                  {importMutation.isPending && (
                    <RefreshCw className="h-4 w-4 animate-spin" />
                  )}{" "}
                  Confirmar importación
                </button>
              </div>
              {!isFacilitator && (
                <div className="text-right text-xs text-amber-300/70">
                  Solo un Facilitador o Administrador puede confirmar una importación.
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {editingRow && (
        <div className="fixed inset-0 z-50 flex justify-end bg-black/70 backdrop-blur-sm">
          <div className="h-full w-full max-w-xl overflow-y-auto border-l border-white/10 bg-[#11191a] shadow-2xl">
            <div className="sticky top-0 z-10 flex items-start justify-between border-b border-white/[0.07] bg-[#11191a]/95 p-5 backdrop-blur-xl">
              <div>
                <div className="font-mono text-xs uppercase tracking-wider text-amber-300">
                  OT {editingRow.workOrder.otNumber}
                </div>
                <h2 className="mt-1 font-display text-2xl text-white">
                  Registrar recepción
                </h2>
                <p className="mt-1 text-sm text-slate-500">
                  {editingRow.equipment?.code} ·{" "}
                  {editingRow.workOrder.description || "Sin descripción"}
                </p>
              </div>
              <button
                onClick={() => setEditingRow(null)}
                className="rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white"
              >
                <X className="h-5 w-5" />
              </button>
            </div>
            <div className="space-y-6 p-5">
              {!editingRequiresLogistics ? (
                <div className="rounded-xl border border-slate-500/20 bg-slate-500/[0.06] p-4 text-sm text-slate-300">
                  <div className="font-semibold text-white">OT sin reserva</div>
                  <p className="mt-1 text-xs text-slate-400">
                    No requiere repuestos ni clasificación logística de recepción.
                  </p>
                </div>
              ) : <>
              <div>
                <div className="mb-3 text-xs font-bold uppercase tracking-[0.18em] text-slate-500">
                  Estado de recepción
                </div>
                <div className="grid gap-2 sm:grid-cols-3">
                  {(["RECIBIDA", "PARCIAL", "FALTA_RECIBIR"] as Status[]).map(
                    status => (
                      <button
                        key={status}
                        onClick={() => setStatusDraft(status)}
                        disabled={!isFacilitator}
                        className={`rounded-xl border p-3 text-left transition ${statusDraft === status ? STATUS_META[status].tone : "border-white/[0.07] bg-white/[0.025] text-slate-500 hover:border-white/15"}`}
                      >
                        <div
                          className={`mb-2 h-2 w-2 rounded-full ${STATUS_META[status].dot}`}
                        />
                        <div className="text-xs font-bold">
                          {STATUS_META[status].label}
                        </div>
                      </button>
                    )
                  )}
                </div>
              </div>
              {statusDraft !== "RECIBIDA" && (
                <div className="rounded-xl border border-amber-300/15 bg-amber-300/[0.05] p-4">
                  <div className="text-xs font-bold uppercase tracking-[0.15em] text-amber-200">
                    ¿Dónde está pendiente?
                  </div>
                  <div className="mt-3 space-y-2">
                    {(bootstrap.data?.warehouses ?? []).map((warehouse: any) => (
                      <label key={warehouse.code} className="flex cursor-pointer items-center gap-3 rounded-lg px-2 py-2 text-sm text-slate-300 hover:bg-white/[0.04]">
                        <input type="checkbox" disabled={!isFacilitator} checked={warehouseCodesDraft.includes(warehouse.code)} onChange={event => setWarehouseCodesDraft(current => event.target.checked ? [...current, warehouse.code] : current.filter(code => code !== warehouse.code))} className="h-4 w-4 accent-amber-300" />
                        {warehouse.name}
                      </label>
                    ))}
                  </div>
                </div>
              )}
              <label className="block text-xs font-bold uppercase tracking-[0.15em] text-slate-500">
                Recibido / revisado por
                  <input
                    value={responsibleDraft}
                    onChange={event => setResponsibleDraft(event.target.value)}
                    disabled={!isFacilitator}
                  className="mt-2 h-11 w-full rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 text-sm normal-case tracking-normal text-white outline-none focus:border-amber-300/40"
                />
              </label>
              <label className="block text-xs font-bold uppercase tracking-[0.15em] text-slate-500">
                Observación
                  <textarea
                    value={noteDraft}
                    onChange={event => setNoteDraft(event.target.value)}
                    disabled={!isFacilitator}
                  rows={4}
                  placeholder="Ej.: pendiente reserva de almacén 2"
                  className="mt-2 w-full resize-none rounded-xl border border-white/[0.08] bg-white/[0.035] p-3 text-sm normal-case tracking-normal text-white outline-none placeholder:text-slate-600 focus:border-amber-300/40"
                />
              </label>
              <div className="rounded-xl border border-white/[0.07] bg-white/[0.025] p-3 text-xs text-slate-500">
                <div className="flex items-center gap-2">
                  <ShieldCheck className="h-4 w-4 text-emerald-300" /> Se
                  guardará usuario, fecha, guardia, estado anterior y nuevo.
                </div>
              </div>
              <div className="flex flex-col gap-2 sm:flex-row">
                <button
                  onClick={submitReception}
                  disabled={
                    !isFacilitator || receptionMutation.isPending ||
                    (statusDraft !== "RECIBIDA" && !warehouseCodesDraft.length)
                  }
                  className="inline-flex flex-1 items-center justify-center gap-2 rounded-xl bg-amber-300 px-4 py-3 text-sm font-bold text-[#0a0f10] disabled:cursor-not-allowed disabled:opacity-40"
                >
                  {receptionMutation.isPending && (
                    <RefreshCw className="h-4 w-4 animate-spin" />
                  )}{" "}
                  Guardar recepción
                </button>
              </div>
              </>}
              <button
                  onClick={() => setShowHistory(!showHistory)}
                  className="inline-flex items-center justify-center gap-2 rounded-xl border border-white/[0.09] px-4 py-3 text-sm font-semibold text-slate-300 hover:bg-white/5"
                >
                  <History className="h-4 w-4" /> Historial
                </button>
              {showHistory && (
                <div className="rounded-xl border border-white/[0.07] bg-black/10 p-4">
                  <div className="mb-3 text-xs font-bold uppercase tracking-[0.15em] text-slate-500">
                    Historial de OT
                  </div>
                  <div className="space-y-3">
                    {history.isLoading ? (
                      <RefreshCw className="h-4 w-4 animate-spin text-amber-300" />
                    ) : history.data?.length ? (
                      history.data.map(item => (
                        <div
                          key={item.id}
                          className="border-l border-amber-300/30 pl-3"
                        >
                          <div className="flex items-center justify-between gap-2">
                            <span className="text-xs font-semibold text-white">
                              {item.newStatus}
                            </span>
                            <span className="text-[10px] text-slate-600">
                              {formatDate(item.createdAt)}
                            </span>
                          </div>
                          <div className="mt-1 text-[11px] text-slate-500">
                            {item.actorName || "Usuario"} · Guardia {item.guard}
                          </div>
                          {item.note && (
                            <div className="mt-1 text-xs text-slate-400">
                              {item.note}
                            </div>
                          )}
                        </div>
                      ))
                    ) : (
                      <div className="text-xs text-slate-600">
                        Sin cambios registrados todavía.
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {showHandover && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-black/75 p-4 backdrop-blur-sm">
          <div className="w-full max-w-lg rounded-2xl border border-white/10 bg-[#11191a] shadow-2xl">
            <div className="flex items-start justify-between border-b border-white/[0.07] p-5">
              <div>
                <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.18em] text-amber-300">
                  <Send className="h-3.5 w-3.5" /> Paquete de relevo
                </div>
                <h2 className="mt-1 font-display text-2xl text-white">
                  {selectedWeek?.label} · Guardia {guard}
                </h2>
              </div>
              <button
                onClick={() => setShowHandover(false)}
                className="rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white"
              >
                <X className="h-5 w-5" />
              </button>
            </div>
            <div className="space-y-5 p-5">
              <div className="grid grid-cols-2 gap-3">
                <div className="rounded-xl border border-white/[0.07] bg-white/[0.025] p-4">
                  <div className="text-[10px] uppercase tracking-wider text-slate-500">
                    Entrega
                  </div>
                  <div className="mt-1 font-display text-2xl text-white">
                    Guardia {guard}
                  </div>
                </div>
                <label className="rounded-xl border border-emerald-400/15 bg-emerald-400/[0.05] p-4">
                  <div className="text-[10px] uppercase tracking-wider text-emerald-300/70">
                    Recibe
                  </div>
                  <select
                    value={handoverTo}
                    onChange={event =>
                      setHandoverTo(event.target.value as Guard)
                    }
                    className="mt-1 bg-transparent font-display text-2xl text-emerald-300 outline-none"
                  >
                    {GUARDS.filter(item => item !== guard).map(item => (
                      <option key={item} value={item} className="bg-[#11191a]">
                        Guardia {item}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
              <div className="rounded-xl border border-rose-400/15 bg-rose-400/[0.05] p-4">
                <div className="flex items-center justify-between">
                  <div>
                    <div className="text-[10px] uppercase tracking-wider text-rose-300/70">
                      Atención requerida
                    </div>
                    <div className="mt-1 font-display text-3xl text-white">
                      {kpis.partial + kpis.missing}{" "}
                      <span className="font-sans text-sm text-slate-500">
                        OT
                      </span>
                    </div>
                  </div>
                  <div className="text-right text-xs text-slate-500">
                    <div>{kpis.missing} falta recibir</div>
                    <div>{kpis.partial} parcial</div>
                  </div>
                </div>
              </div>
              <p className="text-xs leading-5 text-slate-500">
                El paquete guarda la relación de guardias, la fecha, el usuario
                y las OT que requieren atención. SAP continúa siendo la fuente
                oficial.
              </p>
              <div className="flex flex-col gap-2 sm:flex-row sm:justify-end">
                <button
                  onClick={() => setShowHandover(false)}
                  className="rounded-xl px-4 py-2.5 text-sm text-slate-400 hover:bg-white/5"
                >
                  Cancelar
                </button>
                {handoverId ? (
                  <button
                    onClick={() =>
                      confirmHandoverMutation.mutate({ handoverId })
                    }
                    disabled={confirmHandoverMutation.isPending}
                    className="inline-flex items-center justify-center gap-2 rounded-xl bg-emerald-400 px-4 py-2.5 text-sm font-bold text-[#071110] disabled:opacity-40"
                  >
                    <Check className="h-4 w-4" /> Confirmar relevo
                  </button>
                ) : (
                  <button
                    onClick={submitHandover}
                    disabled={handoverMutation.isPending || !handoverTo}
                    className="inline-flex items-center justify-center gap-2 rounded-xl bg-amber-300 px-4 py-2.5 text-sm font-bold text-[#0a0f10] disabled:opacity-40"
                  >
                    {handoverMutation.isPending && (
                      <RefreshCw className="h-4 w-4 animate-spin" />
                    )}{" "}
                    Generar y enviar
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {isAdmin && showAccessManagement && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-black/75 p-4 backdrop-blur-sm">
          <div className="max-h-[90vh] w-full max-w-3xl overflow-y-auto rounded-2xl border border-white/10 bg-[#11191a] shadow-2xl">
            <div className="flex items-start justify-between border-b border-white/[0.07] p-5">
              <div>
                <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.18em] text-amber-300">
                  <ShieldCheck className="h-3.5 w-3.5" /> Administración
                </div>
                <h2 className="mt-1 font-display text-2xl text-white">
                  Gestionar accesos
                </h2>
                <p className="mt-1 text-sm text-slate-500">
                  Asigna el nivel operativo de quienes ya iniciaron sesión en el sistema.
                </p>
              </div>
              <button
                onClick={() => setShowAccessManagement(false)}
                className="rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white"
              >
                <X className="h-5 w-5" />
              </button>
            </div>
            <div className="p-5">
              <div className="overflow-x-auto rounded-xl border border-white/[0.07]">
                <table className="w-full min-w-[620px] text-left text-sm">
                  <thead className="bg-white/[0.035] text-[10px] uppercase tracking-[0.15em] text-slate-500">
                    <tr>
                      <th className="px-4 py-3">Usuario</th>
                      <th className="px-4 py-3">Correo</th>
                      <th className="px-4 py-3">Rol</th>
                      <th className="px-4 py-3">Último acceso</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-white/[0.05]">
                    {accessUsers.isLoading ? (
                      <tr><td colSpan={4} className="px-4 py-8 text-center text-slate-500"><RefreshCw className="mx-auto h-4 w-4 animate-spin text-amber-300" /></td></tr>
                    ) : accessUsers.data?.length ? (
                      accessUsers.data.map(person => (
                        <tr key={person.id}>
                          <td className="px-4 py-3 font-medium text-white">{person.name || "Sin nombre"}</td>
                          <td className="px-4 py-3 text-slate-400">{person.email || "—"}</td>
                          <td className="px-4 py-3">
                            <select
                              value={person.role === "user" ? "supervisor" : person.role}
                              onChange={event => userRoleMutation.mutate({
                                userId: person.id,
                                role: event.target.value as "admin" | "facilitator" | "supervisor" | "viewer",
                              })}
                              disabled={userRoleMutation.isPending}
                              className="h-9 rounded-lg border border-white/[0.08] bg-[#151e20] px-2 text-xs text-white outline-none disabled:opacity-50"
                            >
                              <option value="admin">Administrador</option>
                              <option value="facilitator">Facilitador</option>
                              <option value="supervisor">Supervisor</option>
                              <option value="viewer">Visualizador</option>
                            </select>
                          </td>
                          <td className="px-4 py-3 text-xs text-slate-500">{formatDate(person.lastSignedIn)}</td>
                        </tr>
                      ))
                    ) : (
                      <tr><td colSpan={4} className="px-4 py-8 text-center text-slate-500">Aún no hay usuarios registrados.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
              <p className="mt-3 text-xs leading-5 text-slate-500">
                Administrador: acceso total. Facilitador: importación, recepción y relevo. Supervisor: visualización operativa. Visualizador: solo consulta y KPIs.
              </p>
            </div>
          </div>
        </div>
      )}

      {showRelations && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-black/75 p-4 backdrop-blur-sm">
          <div className="w-full max-w-lg rounded-2xl border border-white/10 bg-[#11191a] shadow-2xl">
            <div className="flex items-start justify-between border-b border-white/[0.07] p-5">
              <div>
                <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.18em] text-amber-300">
                  <Settings2 className="h-3.5 w-3.5" /> Configuración
                </div>
                <h2 className="mt-1 font-display text-2xl text-white">
                  Relaciones de relevo
                </h2>
                <p className="mt-1 text-sm text-slate-500">
                  No se asume una secuencia fija entre guardias.
                </p>
              </div>
              <button
                onClick={() => setShowRelations(false)}
                className="rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white"
              >
                <X className="h-5 w-5" />
              </button>
            </div>
            <div className="space-y-2 p-5">
              {GUARDS.map(from => (
                <div
                  key={from}
                  className="flex items-center gap-3 rounded-xl border border-white/[0.07] bg-white/[0.025] p-3"
                >
                  <div className="grid h-8 w-8 place-items-center rounded-full bg-amber-300/10 font-display text-lg text-amber-200">
                    {from}
                  </div>
                  <ArrowRight className="h-4 w-4 text-slate-600" />
                  <select
                    value={relationsDraft[from]}
                    onChange={event =>
                      setRelationsDraft(prev => ({
                        ...prev,
                        [from]: event.target.value as Guard | "",
                      }))
                    }
                    className="h-10 flex-1 rounded-lg border border-white/[0.08] bg-[#151e20] px-3 text-sm text-white outline-none"
                  >
                    <option value="">Sin relación configurada</option>
                    {GUARDS.filter(to => to !== from).map(to => (
                      <option key={to} value={to}>
                        Guardia {to}
                      </option>
                    ))}
                  </select>
                </div>
              ))}
              <div className="mt-4 flex justify-end gap-2">
                <button
                  onClick={() => setShowRelations(false)}
                  className="rounded-xl px-4 py-2.5 text-sm text-slate-400 hover:bg-white/5"
                >
                  Cancelar
                </button>
                <button
                  onClick={() =>
                    relationsMutation.mutate({
                      relations: GUARDS.filter(
                        from => relationsDraft[from]
                      ).map(from => ({
                        fromGuard: from,
                        toGuard: relationsDraft[from] as Guard,
                      })),
                    })
                  }
                  disabled={relationsMutation.isPending}
                  className="rounded-xl bg-amber-300 px-4 py-2.5 text-sm font-bold text-[#0a0f10] disabled:opacity-40"
                >
                  Guardar relaciones
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
