import { useAuth } from "@/_core/hooks/useAuth";
import { trpc } from "@/lib/trpc";
import {
  mapExcelRows,
  normalizeExcelHeader,
  rowsFromExcelMatrix,
} from "@shared/excelImport";
import { parseCorrectiveLines, type CorrectiveDraft } from "@shared/correctiveImport";
import { canReviewReception, getReservationValue, hasReservation, isCorrectiveWorkOrder } from "@shared/maintenance";
import {
  AlertTriangle,
  ArrowRight,
  BarChart3,
  BellRing,
  Check,
  ChevronDown,
  ClipboardCheck,
  Clock3,
  Eye,
  EyeOff,
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
const GUARD_LABEL: Record<Guard, string> = { A: "1", B: "2", C: "3", D: "4" };
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

function detectWeekFromExcel(fileName: string, matrix: unknown[][]) {
  const source = [fileName, ...matrix.slice(0, 40).flat()]
    .map(value => String(value ?? ""))
    .join(" ");
  const weekMatch = source.match(/(?:sem(?:ana)?|week|wk)[\s_-]*(?:n[°ºo]?\s*)?(\d{1,2})/i);
  if (!weekMatch) return null;
  const weekNumber = Number(weekMatch[1]);
  if (weekNumber < 1 || weekNumber > 53) return null;
  const yearMatch = source.match(/\b(20\d{2})\b/);
  return { weekNumber, year: yearMatch ? Number(yearMatch[1]) : new Date().getFullYear() };
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
  onClick,
}: {
  label: string;
  value: string | number;
  detail?: string;
  color: string;
  icon: typeof PackageCheck;
  onClick?: () => void;
}) {
  return (
    <button type="button" onClick={onClick} className="group relative w-full overflow-hidden rounded-2xl border border-white/[0.07] bg-[#141d1f] p-4 text-left shadow-[0_12px_35px_rgba(0,0,0,0.16)] transition duration-200 hover:-translate-y-0.5 hover:border-white/[0.14]">
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
    </button>
  );
}

export default function Home() {
  const { user, loading: authLoading, isAuthenticated, logout } = useAuth();
  const assignedGuard = ((user as (typeof user & { guard?: Guard }) | null)?.guard) ?? "A";
  const isAdmin = user?.role === "admin";
  const isFacilitator = isAdmin || user?.role === "facilitator";
  const canCloseOt = isAdmin || user?.role === "supervisor";
  const isViewer = user?.role === "viewer";
  const roleLabel = isAdmin
    ? "Administrador"
    : isFacilitator
      ? "Facilitador"
      : isViewer
        ? "Visualizador"
        : "Supervisor";
  const [pin, setPin] = useState("");
  const [selectedLoginUserId, setSelectedLoginUserId] = useState("admin");
  const [showLoginPin, setShowLoginPin] = useState(false);
  const [pinLoginError, setPinLoginError] = useState<string | null>(null);
  const [selectedWeekId, setSelectedWeekId] = useState<number | null>(null);
  const [guard, setGuard] = useState<Guard>("A");
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState<Status | "TODOS">("TODOS");
  const [executionFilter, setExecutionFilter] = useState<"TODAS" | "PENDIENTES" | "CERRADAS">("TODAS");
  const [otTypeFilter, setOtTypeFilter] = useState<"PROGRAMADAS" | "CORRECTIVAS" | "TODAS">("PROGRAMADAS");
  const [warehouseFilter, setWarehouseFilter] = useState("TODOS");
  const [activeSection, setActiveSection] = useState("panel");
  const [historicalEquipmentSearch, setHistoricalEquipmentSearch] = useState("");
  const [historicalTypeFilter, setHistoricalTypeFilter] = useState<"TODAS" | "PROGRAMADAS" | "CORRECTIVAS">("TODAS");
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [showCorrectiveOt, setShowCorrectiveOt] = useState(false);
  const [bulkCorrectiveMode, setBulkCorrectiveMode] = useState(false);
  const [bulkCorrectiveText, setBulkCorrectiveText] = useState("");
  const [bulkCorrectivePreview, setBulkCorrectivePreview] = useState<{ rows: CorrectiveDraft[]; invalid: string[] }>({ rows: [], invalid: [] });
  const [editingCorrective, setEditingCorrective] = useState<{
    id: number;
    otNumber: string;
    equipmentCode: string;
    reservation: string;
    description: string;
  } | null>(null);
  const [correctiveOtNumber, setCorrectiveOtNumber] = useState("");
  const [correctiveEquipment, setCorrectiveEquipment] = useState("");
  const [correctiveReservation, setCorrectiveReservation] = useState("");
  const [correctiveDescription, setCorrectiveDescription] = useState("");
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
  const [equipmentSearch, setEquipmentSearch] = useState("");
  const [newAccessLabel, setNewAccessLabel] = useState("");
  const [newAccessPin, setNewAccessPin] = useState("");
  const [showNewAccessPin, setShowNewAccessPin] = useState(false);
  const [newAccessRole, setNewAccessRole] = useState<"facilitator" | "supervisor">("supervisor");
  const [newAccessGuard, setNewAccessGuard] = useState<Guard>("A");
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
    detectedWeek: { weekNumber: number; year: number } | null;
  } | null>(null);
  const [importWeekNumber, setImportWeekNumber] = useState(
    String(getIsoWeek(new Date()))
  );
  const [importYear, setImportYear] = useState(
    String(new Date().getFullYear())
  );
  const [currentTime, setCurrentTime] = useState(() => new Date());
  const [isDraggingFile, setIsDraggingFile] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const bootstrap = trpc.maintenance.bootstrap.useQuery(undefined, {
    enabled: isAuthenticated,
  });
  const historicalInput = useMemo(() => ({ equipmentQuery: historicalEquipmentSearch.trim(), otType: historicalTypeFilter }), [historicalEquipmentSearch, historicalTypeFilter]);
  const historicalQuery = trpc.maintenance.historicalByEquipment.useQuery(historicalInput, {
    enabled: Boolean(isAuthenticated && (activeSection === "panel" || activeSection === "historico") && historicalInput.equipmentQuery.length > 0),
  });
  const loginUsers = trpc.auth.accessUsers.useQuery();
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
  const accessCodes = trpc.maintenance.listAccessCodes.useQuery(undefined, {
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
  const createCorrectiveOtMutation = trpc.maintenance.createCorrectiveOt.useMutation({
    onSuccess: async () => {
      toast.success("OT correctiva creada y visible en el panel.");
      setShowCorrectiveOt(false);
      setOtTypeFilter("CORRECTIVAS");
      setCorrectiveOtNumber("");
      setCorrectiveEquipment("");
      setCorrectiveReservation("");
      setCorrectiveDescription("");
      await utils.maintenance.dashboard.invalidate();
    },
    onError: error => toast.error(error.message),
  });
  const createCorrectiveOtsMutation = trpc.maintenance.createCorrectiveOts.useMutation({
    onSuccess: async result => {
      toast.success(`${result.created} OT correctivas creadas${result.duplicates ? ` · ${result.duplicates} repetidas omitidas` : ""}.`);
      setShowCorrectiveOt(false);
      setBulkCorrectiveMode(false);
      setOtTypeFilter("CORRECTIVAS");
      setBulkCorrectiveText("");
      setBulkCorrectivePreview({ rows: [], invalid: [] });
      await utils.maintenance.dashboard.invalidate();
    },
    onError: error => toast.error(error.message),
  });
  const updateCorrectiveOtMutation = trpc.maintenance.updateCorrectiveOt.useMutation({
    onSuccess: () => {
      toast.success("OT correctiva actualizada.");
      setEditingCorrective(null);
      void utils.maintenance.dashboard.invalidate();
    },
    onError: error => toast.error(error.message),
  });
  const cancelCorrectiveOtMutation = trpc.maintenance.cancelCorrectiveOt.useMutation({
    onSuccess: () => {
      toast.success("OT correctiva anulada y conservada en el historial.");
      void utils.maintenance.dashboard.invalidate();
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
  const createAccessCodeMutation = trpc.maintenance.createAccessCode.useMutation({
    onSuccess: () => {
      toast.success("Usuario y PIN creados.");
      setNewAccessLabel("");
      setNewAccessPin("");
      void accessCodes.refetch();
    },
    onError: error => toast.error(error.message),
  });
  const revokeAccessCodeMutation = trpc.maintenance.revokeAccessCode.useMutation({
    onSuccess: () => {
      toast.success("Acceso desactivado.");
      void accessCodes.refetch();
    },
    onError: error => toast.error(error.message),
  });
  const updateAccessCodeGuardMutation = trpc.maintenance.updateAccessCodeGuard.useMutation({
    onSuccess: () => {
      toast.success("Guardia del usuario actualizada.");
      void accessCodes.refetch();
    },
    onError: error => toast.error(error.message),
  });

  const weeks = bootstrap.data?.weeks ?? [];
  const allEquipment =
    dashboard.data?.allEquipment ?? bootstrap.data?.equipment ?? [];
  const filteredEquipment = useMemo(() => {
    const query = equipmentSearch.trim().toLowerCase();
    return allEquipment.filter(item =>
      !query || `${item.code} ${item.description ?? ""}`.toLowerCase().includes(query)
    );
  }, [allEquipment, equipmentSearch]);
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
    const hasDescriptionColumn = selected.some(key => {
      const normalized = normalizeExcelHeader(key);
      return normalized === "descripcion" || normalized.includes("descripciondeoperacion") || normalized.includes("descripcionoperacion");
    });
    if (!hasDescriptionColumn) selected.push("Descripción de operación");
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
    if (!isAdmin) setGuard(assignedGuard);
  }, [assignedGuard, isAdmin]);

  useEffect(() => {
    const timer = window.setInterval(() => setCurrentTime(new Date()), 1000);
    return () => window.clearInterval(timer);
  }, []);

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
      const matchesExecution =
        executionFilter === "TODAS" ||
        (executionFilter === "CERRADAS" && row.execution?.status === "CERRADA_NOTIFICADA") ||
        (executionFilter === "PENDIENTES" && row.execution?.status !== "CERRADA_NOTIFICADA");
      const isCorrective = isCorrectiveWorkOrder(row.workOrder.sourceData);
      const matchesOtType =
        otTypeFilter === "TODAS" ||
        (otTypeFilter === "CORRECTIVAS" && isCorrective) ||
        (otTypeFilter === "PROGRAMADAS" && !isCorrective);
      return matchesSearch && matchesStatus && matchesWarehouse && matchesExecution && matchesOtType;
    });
  }, [rows, search, statusFilter, warehouseFilter, executionFilter, otTypeFilter]);

  const otTypeCounts = useMemo(() => rows.reduce(
    (counts, row) => {
      const isCorrective = isCorrectiveWorkOrder(row.workOrder.sourceData);
      if (isCorrective) counts.correctivas += 1;
      else counts.programadas += 1;
      return counts;
    },
    { programadas: 0, correctivas: 0 }
  ), [rows]);

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
    setResponsibleDraft(user?.name ?? "Usuario actual");
    setNoteDraft(row.state?.note ?? "");
    setShowHistory(false);
  }

  function openCorrectiveEditor(row: any) {
    const fields = getImportedFields(row.workOrder.sourceData);
    setEditingCorrective({
      id: row.workOrder.id,
      otNumber: row.workOrder.otNumber,
      equipmentCode: row.equipment?.code ?? "",
      reservation: fields.Reserva ?? "",
      description: row.workOrder.description ?? fields["Descripción de operación"] ?? "",
    });
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
      const detectedWeek = detectWeekFromExcel(file.name, matrix);
      if (detectedWeek) {
        setImportWeekNumber(String(detectedWeek.weekNumber));
        setImportYear(String(detectedWeek.year));
      }
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
        detectedWeek,
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

  function setFilteredEquipmentSelected(selected: boolean) {
    if (!isFacilitator || !selectedWeekId) return;
    const visibleIds = filteredEquipment
      .map(item => item.id)
      .filter(id => Number.isSafeInteger(id) && id > 0);
    const next = selected
      ? Array.from(new Set([...selectedEquipmentIds, ...visibleIds]))
      : selectedEquipmentIds.filter(id => !visibleIds.includes(id));
    setSelectedEquipmentIds(next);
    assignmentMutation.mutate({ weekId: selectedWeekId, guard, equipmentIds: next });
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
                trazabilidad. El Administrador crea un PIN para cada Facilitador
                o Supervisor; no se requiere correo.
              </p>
              <form
                className="mt-9 flex max-w-xl flex-col gap-3 sm:flex-row"
                onSubmit={event => {
                  event.preventDefault();
                  setPinLoginError(null);
                  pinLoginMutation.mutate({ pin, selectedUserId: selectedLoginUserId });
                }}
              >
                <select
                  aria-label="Seleccionar usuario"
                  value={selectedLoginUserId}
                  onChange={event => setSelectedLoginUserId(event.target.value)}
                  className="h-12 w-full rounded-xl border border-white/10 bg-[#121a1c] px-4 text-sm text-white outline-none transition focus:border-amber-300/60 focus:ring-2 focus:ring-amber-300/20 sm:w-64"
                >
                  {(loginUsers.data ?? [{ id: "admin", label: "Administrador", role: "admin", guard: "A" }]).map(loginUser => (
                    <option key={loginUser.id} value={loginUser.id}>
                      {loginUser.label}{loginUser.role === "admin" ? " · Administrador" : ` · Guardia ${loginUser.guard === "A" ? "1" : loginUser.guard === "B" ? "2" : loginUser.guard === "C" ? "3" : "4"}`}
                    </option>
                  ))}
                </select>
                <div className="relative min-w-0 flex-1">
                  <input
                    aria-label="PIN o código de acceso"
                    autoComplete="current-password"
                    className="w-full rounded-xl border border-white/10 bg-[#121a1c] px-4 py-3 pr-12 text-white outline-none transition placeholder:text-slate-500 focus:border-amber-300/60 focus:ring-2 focus:ring-amber-300/20 disabled:opacity-60"
                    maxLength={128}
                    minLength={8}
                    onChange={event => setPin(event.target.value)}
                    placeholder="PIN o código de acceso"
                    required
                    type={showLoginPin ? "text" : "password"}
                    value={pin}
                  />
                  <button
                    type="button"
                    aria-label={showLoginPin ? "Ocultar PIN" : "Mostrar PIN"}
                    title={showLoginPin ? "Ocultar PIN" : "Mostrar PIN"}
                    onClick={() => setShowLoginPin(current => !current)}
                    className="absolute right-2 top-1/2 -translate-y-1/2 rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-amber-200"
                  >
                    {showLoginPin ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                  </button>
                </div>
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
                      SEM 38 · GUARDIA 1
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
          className={`fixed inset-y-0 left-0 z-40 overflow-hidden border-r border-white/[0.07] bg-[#0d1415] px-4 py-5 transition-all lg:sticky lg:top-0 lg:block lg:h-screen lg:translate-x-0 ${sidebarCollapsed ? "lg:w-[78px] lg:px-2" : "lg:w-[270px] lg:px-4"} ${mobileMenuOpen ? "translate-x-0 w-[270px]" : "-translate-x-full w-[270px]"}`}
        >
          <div className={`relative flex items-center px-2 ${sidebarCollapsed ? "lg:justify-center" : "lg:justify-between"}`}>
            <div className="flex items-center gap-3">
              <div className="grid h-10 w-10 place-items-center rounded-xl bg-amber-300 text-[#0a0f10]">
                <HardHat className="h-5 w-5" />
              </div>
              <div className={sidebarCollapsed ? "lg:hidden" : ""}>
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
            <button
              className="hidden rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white lg:absolute lg:right-0 lg:block"
              onClick={() => setSidebarCollapsed(current => !current)}
              aria-label={sidebarCollapsed ? "Mostrar panel lateral" : "Ocultar panel lateral"}
              title={sidebarCollapsed ? "Mostrar panel lateral" : "Ocultar panel lateral"}
            >
              <PanelLeftClose className={`h-4 w-4 transition-transform ${sidebarCollapsed ? "rotate-180" : ""}`} />
            </button>
          </div>
          <div className={`mt-10 px-2 text-[10px] font-bold uppercase tracking-[0.2em] text-slate-600 ${sidebarCollapsed ? "lg:hidden" : ""}`}>
            Operación
          </div>
          <nav className="mt-3 space-y-1">
            {[
              { id: "panel", label: "Panel de control", icon: BarChart3 },
              { id: "equipos", label: "Equipos de mi guardia", icon: Truck },
              { id: "ot", label: "OT y recepción", icon: ClipboardCheck },
              { id: "historico", label: "Histórico de OT", icon: History },
            ].filter(item => item.id !== "relevo" || isFacilitator).map(item => (
              <button
                key={item.id}
                onClick={() => scrollTo(item.id)}
                className={`flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left text-sm transition ${sidebarCollapsed ? "lg:justify-center lg:px-2" : ""} ${activeSection === item.id ? "bg-amber-300/10 font-semibold text-amber-200" : "text-slate-400 hover:bg-white/[0.04] hover:text-white"}`}
              >
                <item.icon className="h-4 w-4" />
                <span className={sidebarCollapsed ? "lg:hidden" : ""}>{item.label}</span>
              </button>
            ))}
          </nav>
          <div className={`mt-9 px-2 text-[10px] font-bold uppercase tracking-[0.2em] text-slate-600 ${sidebarCollapsed ? "lg:hidden" : ""}`}>
            Administración
          </div>
          <nav className="mt-3 space-y-1">
            {isFacilitator && (
              <button
                onClick={() => setShowImport(true)}
                className={`flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left text-sm text-slate-400 transition hover:bg-white/[0.04] hover:text-white ${sidebarCollapsed ? "lg:justify-center lg:px-2" : ""}`}
              >
                <Import className="h-4 w-4" />
                <span className={sidebarCollapsed ? "lg:hidden" : ""}>Importar Excel semanal</span>
              </button>
            )}
            {isAdmin && (
              <button
                onClick={() => setShowAccessManagement(true)}
                className={`flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left text-sm text-slate-400 transition hover:bg-white/[0.04] hover:text-white ${sidebarCollapsed ? "lg:justify-center lg:px-2" : ""}`}
              >
                <ShieldCheck className="h-4 w-4" />
                <span className={sidebarCollapsed ? "lg:hidden" : ""}>Gestionar accesos</span>
              </button>
            )}
          </nav>
          <div className={`absolute bottom-5 rounded-2xl border border-white/[0.07] bg-white/[0.025] p-3 ${sidebarCollapsed ? "left-2 right-2 lg:p-2" : "left-4 right-4"}`}>
            <div className={`flex items-center gap-3 ${sidebarCollapsed ? "lg:justify-center" : ""}`}>
              <div className="grid h-9 w-9 place-items-center rounded-full bg-slate-700 text-sm font-bold text-white">
                {(user?.name ?? "U").slice(0, 1).toUpperCase()}
              </div>
              <div className={`min-w-0 flex-1 ${sidebarCollapsed ? "lg:hidden" : ""}`}>
                <div className="truncate text-sm font-semibold text-white">
                  {user?.name ?? "Usuario"}
                </div>
                <div className="text-[10px] uppercase tracking-wider text-slate-500">
                  {roleLabel}{!isAdmin && ` · Guardia ${GUARD_LABEL[assignedGuard]}`}
                </div>
              </div>
              <button
                className={sidebarCollapsed ? "lg:hidden" : "rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white"}
                onClick={() => logout()}
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
                <div className="flex flex-wrap items-center justify-end gap-2">
                  <div className="mr-1 flex items-center gap-2 rounded-xl border border-white/[0.07] bg-white/[0.025] px-3 py-2 text-right">
                    <Clock3 className="h-4 w-4 text-amber-300" />
                    <div>
                      <div className="font-mono text-sm font-bold tabular-nums text-white">
                        {currentTime.toLocaleTimeString("es-PE", { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
                      </div>
                      <div className="text-[9px] uppercase tracking-[0.12em] text-slate-500">
                        {currentTime.toLocaleDateString("es-PE", { weekday: "short", day: "2-digit", month: "short", year: "numeric" })}
                      </div>
                    </div>
                  </div>
                  {isFacilitator && (
                    <button
                      onClick={() => setShowImport(true)}
                      className="inline-flex items-center gap-2 rounded-xl border border-amber-300/25 bg-amber-300/10 px-3.5 py-2.5 text-xs font-bold text-amber-200 transition hover:bg-amber-300/15"
                    >
                      <Upload className="h-4 w-4" /> Importar Excel
                    </button>
                  )}
                  {isFacilitator && (
                    <button
                      onClick={() => setShowCorrectiveOt(true)}
                      className="inline-flex items-center gap-2 rounded-xl border border-sky-300/25 bg-sky-300/10 px-3.5 py-2.5 text-xs font-bold text-sky-200 transition hover:bg-sky-300/15"
                    >
                      <ClipboardCheck className="h-4 w-4" /> Crear OT correctiva
                    </button>
                  )}
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
                    {GUARDS.filter(item => isAdmin || item === assignedGuard).map(item => (
                      <button
                        key={item}
                        onClick={() => setGuard(item)}
                        className={`rounded-lg px-3 py-1.5 text-xs font-bold transition ${guard === item ? "bg-amber-300 text-[#0a0f10]" : "text-slate-500 hover:text-white"}`}
                      >
                        Guardia {GUARD_LABEL[item]}
                      </button>
                    ))}
                  </div>
                </div>
                <div className="flex items-center gap-2 text-xs text-slate-500">
                  <span>Guardia actual</span>
                  <span className="rounded-full border border-white/10 px-2 py-1 font-semibold text-slate-300">
                    {GUARD_LABEL[guard]}
                  </span>
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
                    onClick={() => scrollTo("ot")}
                  />
                  <KpiCard
                    label="Recibidas"
                    value={kpis.received}
                    detail={`${kpis.percent}% completas`}
                    color="bg-emerald-400"
                    icon={PackageCheck}
                    onClick={() => { setStatusFilter("RECIBIDA"); scrollTo("ot"); }}
                  />
                  <KpiCard
                    label="Parciales"
                    value={kpis.partial}
                    detail="Requieren seguimiento"
                    color="bg-amber-400"
                    icon={AlertTriangle}
                    onClick={() => { setStatusFilter("PARCIAL"); scrollTo("ot"); }}
                  />
                  <KpiCard
                    label="Falta recibir"
                    value={kpis.missing}
                    detail="Atención prioritaria"
                    color="bg-rose-400"
                    icon={BellRing}
                    onClick={() => { setStatusFilter("FALTA_RECIBIR"); scrollTo("ot"); }}
                  />
                  <KpiCard
                    label="Recepción total"
                    value={`${kpis.percent}%`}
                    detail="Solo OT completas"
                    color="bg-violet-400"
                    icon={BarChart3}
                    onClick={() => scrollTo("ot")}
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
                        la Guardia {GUARD_LABEL[guard]}.
                      </p>
                    </div>
                    <div className="rounded-xl border border-amber-300/20 bg-amber-300/10 px-3 py-2 text-xs text-amber-200">
                      {selectedEquipmentIds.length} seleccionados
                    </div>
                  </div>
                  <div className="mt-5 grid grid-cols-2 gap-2 sm:grid-cols-4 lg:grid-cols-6">
                    <div className="col-span-full flex flex-col gap-2 rounded-xl border border-white/[0.07] bg-white/[0.02] p-3 sm:flex-row sm:items-center">
                      <div className="flex flex-1 items-center gap-2 rounded-lg border border-white/[0.08] bg-[#0d1314] px-3 py-2">
                        <Search className="h-4 w-4 text-slate-500" />
                        <input
                          value={equipmentSearch}
                          onChange={event => setEquipmentSearch(event.target.value)}
                          placeholder="Buscar equipo por código o descripción..."
                          className="w-full bg-transparent text-sm text-white outline-none placeholder:text-slate-600"
                        />
                      </div>
                      {isFacilitator && (
                        <div className="flex gap-2">
                          <button
                            onClick={() => setFilteredEquipmentSelected(true)}
                            disabled={!filteredEquipment.length || assignmentMutation.isPending}
                            className="rounded-lg border border-emerald-400/20 px-3 py-2 text-xs font-semibold text-emerald-300 hover:bg-emerald-400/10 disabled:opacity-40"
                          >
                            Seleccionar visibles
                          </button>
                          <button
                            onClick={() => setFilteredEquipmentSelected(false)}
                            disabled={!filteredEquipment.length || assignmentMutation.isPending}
                            className="rounded-lg border border-rose-400/20 px-3 py-2 text-xs font-semibold text-rose-300 hover:bg-rose-400/10 disabled:opacity-40"
                          >
                            Deseleccionar visibles
                          </button>
                        </div>
                      )}
                    </div>
                    {filteredEquipment.map(item => (
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
                    {!filteredEquipment.length && allEquipment.length > 0 && (
                      <div className="col-span-full rounded-xl border border-dashed border-white/10 p-6 text-center text-sm text-slate-500">
                        No hay equipos que coincidan con la búsqueda.
                      </div>
                    )}
                    {!allEquipment.length && (
                      <div className="col-span-full rounded-xl border border-dashed border-white/10 p-6 text-center text-sm text-slate-500">
                        Importa una semana para cargar equipos.
                      </div>
                    )}
                  </div>
                </section>

                <section
                  id="ot"
                  className="mt-6 scroll-mt-24 rounded-2xl border border-white/[0.07] bg-[#101718] p-4"
                >
                  <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
                    <div>
                      <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.18em] text-amber-300">
                        <ClipboardCheck className="h-3.5 w-3.5" /> Control
                        operativo
                      </div>
                      <h2 className="mt-1 font-display text-2xl text-white">
                        {otTypeFilter === "CORRECTIVAS" ? "OT correctivas" : otTypeFilter === "PROGRAMADAS" ? "OT programadas" : "Todas las OT"} · Guardia {GUARD_LABEL[guard]}
                      </h2>
                      <p className="mt-1 text-xs text-slate-500">
                        Consulta cada tipo de OT por separado, con su equipo, reserva y descripción de operación.
                      </p>
                    </div>
                    <div className="flex flex-col gap-2 sm:flex-row">
                      <div className="flex rounded-xl border border-white/[0.08] bg-white/[0.025] p-1" role="tablist" aria-label="Tipo de OT">
                        {([
                          ["PROGRAMADAS", "Programadas", otTypeCounts.programadas],
                          ["CORRECTIVAS", "Correctivas", otTypeCounts.correctivas],
                          ["TODAS", "Todas", otTypeCounts.programadas + otTypeCounts.correctivas],
                        ] as const).map(([value, label, count]) => (
                          <button
                            key={value}
                            type="button"
                            role="tab"
                            aria-selected={otTypeFilter === value}
                            onClick={() => setOtTypeFilter(value)}
                            className={`rounded-lg px-3 py-2 text-[11px] font-bold transition ${otTypeFilter === value ? "bg-amber-300 text-[#0a0f10]" : "text-slate-400 hover:bg-white/[0.05] hover:text-white"}`}
                          >
                            {label} <span className="ml-1 opacity-70">{count}</span>
                          </button>
                        ))}
                      </div>
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
                      {canCloseOt && (
                        <select
                          value={executionFilter}
                          onChange={event => setExecutionFilter(event.target.value as typeof executionFilter)}
                          className="h-10 rounded-xl border border-blue-300/15 bg-[#151e20] px-3 text-xs text-slate-300 outline-none"
                          aria-label="Filtrar ejecución de OT"
                        >
                          <option value="TODAS">Todas las OT</option>
                          <option value="PENDIENTES">Pendientes de cierre</option>
                          <option value="CERRADAS">OT cerradas</option>
                        </select>
                      )}
                    </div>
                  </div>
                  <div className="mt-3 flex flex-wrap items-center gap-x-4 gap-y-2 rounded-xl border border-white/[0.06] bg-white/[0.02] px-3 py-2 text-[10px] font-semibold uppercase tracking-[0.12em] text-slate-500">
                    <span className="text-slate-400">Estados visuales:</span>
                    <span className="inline-flex items-center gap-1.5"><i className="h-2.5 w-2.5 rounded-full bg-emerald-400" /> Recibida</span>
                    <span className="inline-flex items-center gap-1.5"><i className="h-2.5 w-2.5 rounded-full bg-orange-400" /> Parcial / pendiente</span>
                    <span className="inline-flex items-center gap-1.5"><i className="h-2.5 w-2.5 rounded-full bg-blue-400" /> Cerrada</span>
                  </div>
                  <div className="mt-4 max-h-[calc(100vh-360px)] min-h-[260px] overflow-auto rounded-xl border border-white/[0.06]">
                    <table className="w-full min-w-max text-left">
                      <thead className="sticky top-0 z-10 bg-[#162022] text-[10px] uppercase tracking-[0.15em] text-slate-600">
                        <tr>
                          <th className="whitespace-nowrap px-4 py-3 font-bold">
                            OT
                          </th>
                          <th className="whitespace-nowrap px-4 py-3 font-bold">
                            Equipo
                          </th>
                          <th className="whitespace-nowrap px-3 py-3 font-bold">
                            Acción
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
                            Responsable / cierre
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
                          const fields = getImportedFields(
                            row.workOrder.sourceData
                          );
                          const isCorrective = fields.Tipo === "Correctiva";
                          const requiresLogistics =
                            row.requiresLogistics ?? hasReservation(row.workOrder.sourceData);
                          const rowTone = row.execution?.status === "CERRADA_NOTIFICADA"
                            ? "bg-blue-500/[0.12] hover:bg-blue-500/[0.18]"
                            : state === "RECIBIDA"
                              ? "bg-emerald-400/[0.10] hover:bg-emerald-400/[0.16]"
                              : state === "PARCIAL" || state === "FALTA_RECIBIR"
                                ? "bg-orange-400/[0.10] hover:bg-orange-400/[0.16]"
                                : "hover:bg-white/[0.025]";
                          return (
                            <tr
                              key={row.workOrder.id}
                              className={`group border-l-2 transition ${rowTone} ${row.execution?.status === "CERRADA_NOTIFICADA" ? "border-l-blue-400" : state === "RECIBIDA" ? "border-l-emerald-400" : state === "PARCIAL" || state === "FALTA_RECIBIR" ? "border-l-orange-400" : "border-l-transparent"}`}
                            >
                              <td className="whitespace-nowrap px-4 py-2.5 font-mono text-sm font-semibold text-white">
                                {row.workOrder.otNumber}
                              </td>
                              <td className="whitespace-nowrap px-4 py-2.5 text-sm text-amber-200/85">
                                {row.equipment?.code ?? "—"}
                              </td>
                              <td className="whitespace-nowrap px-3 py-2.5">
                                {isFacilitator && requiresLogistics && (
                                  !canReviewReception(row.state?.status as Status | undefined) ? (
                                    <span className="inline-flex items-center rounded-xl border border-emerald-300/20 bg-emerald-300/[0.08] px-3.5 py-2 text-xs font-bold text-emerald-200">
                                      Revisada
                                    </span>
                                  ) : (
                                    <button
                                      onClick={() => openEdit(row)}
                                      className="rounded-xl border border-amber-300/35 bg-amber-300/10 px-3.5 py-2 text-xs font-bold text-amber-200 shadow-sm transition hover:bg-amber-300/20 hover:text-amber-100"
                                    >
                                      Revisar
                                    </button>
                                  )
                                )}
                              </td>
                              {excelColumns.map(column => {
                                const normalizedColumn = normalizeExcelHeader(column);
                                const isReservationColumn = normalizedColumn.includes("reserva");
                                const isDescriptionColumn = normalizedColumn === "descripcion" || normalizedColumn.includes("descripciondeoperacion") || normalizedColumn.includes("descripcionoperacion");
                                const value = formatImportedCell(
                                  fields[column] ||
                                    (isReservationColumn
                                      ? getReservationValue(row.workOrder.sourceData)
                                      : isDescriptionColumn
                                        ? row.workOrder.description ?? ""
                                        : undefined)
                                );
                                return (
                                  <td
                                    key={column}
                                    className={`${isDescriptionColumn ? "min-w-[320px] max-w-[420px] whitespace-normal" : "max-w-[190px]"} px-3 py-2.5 text-xs text-slate-400`}
                                  >
                                    <div className={isDescriptionColumn ? "whitespace-normal break-words leading-relaxed" : "truncate"} title={value}>
                                      {value}
                                    </div>
                                  </td>
                                );
                              })}
                              <td className="whitespace-nowrap px-4 py-2.5">
                                {requiresLogistics ? (
                                  <StatusPill status={state} />
                                ) : (
                                  <span className="text-xs text-slate-500">
                                    Sin reserva
                                  </span>
                                )}
                              </td>
                              <td className="min-w-[190px] px-4 py-2.5 text-xs text-slate-400">
                                <div className="space-y-1.5">
                                  <div className="flex items-center justify-between gap-3">
                                    <span className="text-[9px] font-bold uppercase tracking-[0.12em] text-slate-600">Recepción</span>
                                    <span className="truncate font-semibold text-slate-300">{row.state?.responsibleName || "Sin registrar"}</span>
                                  </div>
                                  {row.execution?.status === "CERRADA_NOTIFICADA" && (
                                    <div className="flex items-center justify-between gap-3 border-t border-white/[0.06] pt-1.5">
                                      <span className="text-[9px] font-bold uppercase tracking-[0.12em] text-blue-300/60">Cierre</span>
                                      <span className="truncate text-[10px] font-semibold text-blue-200" title={`${row.execution.executedByName || "Usuario"} · ${formatDate(row.execution.executedAt)}`}>
                                        {row.execution.executedByName || "Usuario"} · {formatDate(row.execution.executedAt)}
                                      </span>
                                    </div>
                                  )}
                                </div>
                              </td>
                              <td className="whitespace-nowrap px-4 py-2.5 text-xs text-slate-500">
                                {formatDate(row.state?.updatedAt)}
                              </td>
                              <td className="whitespace-nowrap px-4 py-2.5 text-right">
                                {canCloseOt && selectedWeekId && (
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
                                {isFacilitator && isCorrective && (
                                  <>
                                    <button
                                      onClick={() => openCorrectiveEditor(row)}
                                      className="ml-2 rounded-lg border border-sky-300/20 px-2.5 py-1.5 text-[11px] font-semibold text-sky-200 transition hover:bg-sky-300/10"
                                    >
                                      Editar
                                    </button>
                                    <button
                                      onClick={() => {
                                        if (window.confirm("¿Anular esta OT correctiva? Se conservará en el historial, pero dejará de aparecer en el panel activo."))
                                          cancelCorrectiveOtMutation.mutate({ workOrderId: row.workOrder.id });
                                      }}
                                      disabled={cancelCorrectiveOtMutation.isPending}
                                      className="ml-2 rounded-lg border border-rose-300/20 px-2.5 py-1.5 text-[11px] font-semibold text-rose-200 transition hover:bg-rose-300/10 disabled:opacity-40"
                                    >
                                      Anular
                                    </button>
                                  </>
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

                <section id="historico" className="mt-6 scroll-mt-24 rounded-2xl border border-white/[0.07] bg-[#101718] p-5">
                  <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
                    <div>
                      <div className="flex items-center gap-2 text-[10px] font-bold uppercase tracking-[0.18em] text-violet-300">
                        <History className="h-3.5 w-3.5" /> Consulta histórica
                      </div>
                      <h2 className="mt-1 font-display text-2xl text-white">Histórico de OT por equipo</h2>
                      <p className="mt-1 text-sm text-slate-500">Busca un equipo para consultar sus OT de todas las semanas registradas.</p>
                    </div>
                    <div className="flex w-full flex-col gap-2 sm:w-auto sm:flex-row">
                      <div className="relative w-full sm:w-80">
                        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-600" />
                        <input
                          value={historicalEquipmentSearch}
                          onChange={event => setHistoricalEquipmentSearch(event.target.value)}
                          placeholder="Ej. CA05, CR02..."
                          className="h-11 w-full rounded-xl border border-violet-300/20 bg-white/[0.035] pl-9 pr-3 text-sm text-white outline-none placeholder:text-slate-600 focus:border-violet-300/50"
                        />
                      </div>
                      <select
                        value={historicalTypeFilter}
                        onChange={event => setHistoricalTypeFilter(event.target.value as typeof historicalTypeFilter)}
                        className="h-11 rounded-xl border border-violet-300/20 bg-[#151e20] px-3 text-xs text-slate-300 outline-none focus:border-violet-300/50"
                        aria-label="Filtrar histórico por tipo de OT"
                      >
                        <option value="TODAS">Todas las OT</option>
                        <option value="PROGRAMADAS">Programadas</option>
                        <option value="CORRECTIVAS">Correctivas</option>
                      </select>
                    </div>
                  </div>
                  {!historicalInput.equipmentQuery ? (
                    <div className="mt-5 rounded-xl border border-dashed border-white/10 p-8 text-center text-sm text-slate-500">Escribe un código de equipo para consultar su histórico.</div>
                  ) : historicalQuery.isLoading ? (
                    <div className="mt-5 rounded-xl border border-white/[0.06] p-8 text-center text-sm text-slate-500">Buscando OT…</div>
                  ) : historicalQuery.data?.length ? (
                    <div className="mt-5 overflow-x-auto rounded-xl border border-white/[0.06]">
                      <table className="w-full min-w-[760px] text-left">
                        <thead className="bg-white/[0.035] text-[10px] uppercase tracking-[0.15em] text-slate-600">
                          <tr>
                            <th className="px-4 py-3 font-bold">Semana</th>
                            <th className="px-4 py-3 font-bold">OT</th>
                            <th className="px-4 py-3 font-bold">Equipo</th>
                            <th className="px-4 py-3 font-bold">Reserva</th>
                            <th className="px-4 py-3 font-bold">Descripción</th>
                            <th className="px-4 py-3 font-bold">Recepción</th>
                            <th className="px-4 py-3 font-bold">Actualizado</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-white/[0.05]">
                          {historicalQuery.data.map(row => {
                            const fields = getImportedFields(row.workOrder.sourceData);
                            const historicalStatus = (row.state?.status ?? "SIN_REVISAR") as Status;
                            return (
                              <tr key={row.workOrder.id} className="hover:bg-white/[0.025]">
                                <td className="whitespace-nowrap px-4 py-3 text-xs font-semibold text-violet-200">SEM {row.week.weekNumber} · {row.week.year}</td>
                                <td className="whitespace-nowrap px-4 py-3 font-mono text-sm font-semibold text-white">{row.workOrder.otNumber}</td>
                                <td className="whitespace-nowrap px-4 py-3 font-mono text-sm text-amber-200">{row.equipment.code}</td>
                                <td className="whitespace-nowrap px-4 py-3 text-xs text-slate-400">{getReservationValue(row.workOrder.sourceData) || "Sin reserva"}</td>
                                <td className="max-w-[320px] px-4 py-3 text-xs text-slate-300" title={row.workOrder.description ?? fields["Descripción de operación"] ?? ""}>{row.workOrder.description ?? fields["Descripción de operación"] ?? "—"}</td>
                                <td className="whitespace-nowrap px-4 py-3">{row.state ? <StatusPill status={historicalStatus} /> : <span className="text-xs text-slate-600">Sin revisar</span>}</td>
                                <td className="whitespace-nowrap px-4 py-3 text-xs text-slate-500">{formatDate(row.state?.updatedAt)}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div className="mt-5 rounded-xl border border-dashed border-white/10 p-8 text-center text-sm text-slate-500">No se encontraron OT para ese equipo.</div>
                  )}
                </section>

                {false && isFacilitator && <section
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
                        Guardia {GUARD_LABEL[guard]}
                      </div>
                    </div>
                    <div className="rounded-xl border border-white/[0.07] bg-black/10 p-3">
                      <div className="text-[10px] uppercase tracking-widest text-slate-500">
                        Recibe
                      </div>
                      <div className="mt-1 font-display text-xl text-emerald-300">
                        Guardia {selectedRelation ? GUARD_LABEL[selectedRelation as Guard] : "—"}
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

      {showCorrectiveOt && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-black/75 p-4 backdrop-blur-sm">
          <div className="w-full max-w-xl rounded-2xl border border-sky-300/20 bg-[#11191a] shadow-2xl">
            <div className="flex items-start justify-between border-b border-white/[0.07] p-5">
              <div>
                <div className="text-[10px] font-bold uppercase tracking-[0.18em] text-sky-300">Nueva orden</div>
                <h2 className="mt-1 font-display text-2xl text-white">Crear OT correctiva</h2>
                <p className="mt-1 text-sm text-slate-500">Se guardará en Semana {selectedWeek?.weekNumber ?? "—"} · Guardia {GUARD_LABEL[guard]} y quedará disponible para revisión.</p>
              </div>
              <button onClick={() => setShowCorrectiveOt(false)} className="rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white"><X className="h-5 w-5" /></button>
            </div>
            <form
              className="grid gap-4 p-5"
              onSubmit={event => {
                event.preventDefault();
                if (!selectedWeekId) return toast.error("Selecciona una semana antes de crear la OT.");
                if (bulkCorrectiveMode) {
                  if (!bulkCorrectivePreview.rows.length) return toast.error("Pega al menos una línea válida.");
                  createCorrectiveOtsMutation.mutate({ weekId: selectedWeekId, guard, rows: bulkCorrectivePreview.rows });
                } else {
                  createCorrectiveOtMutation.mutate({ weekId: selectedWeekId, guard, otNumber: correctiveOtNumber, equipmentCode: correctiveEquipment, reservation: correctiveReservation || undefined, description: correctiveDescription });
                }
              }}
            >
              <div className="flex rounded-xl bg-white/[0.04] p-1">
                <button type="button" onClick={() => setBulkCorrectiveMode(false)} className={`flex-1 rounded-lg px-3 py-2 text-xs font-bold ${!bulkCorrectiveMode ? "bg-sky-300 text-[#071014]" : "text-slate-400"}`}>Una OT</button>
                <button type="button" onClick={() => setBulkCorrectiveMode(true)} className={`flex-1 rounded-lg px-3 py-2 text-xs font-bold ${bulkCorrectiveMode ? "bg-sky-300 text-[#071014]" : "text-slate-400"}`}>Pegar varias OT</button>
              </div>
              {!bulkCorrectiveMode ? <>
                <div className="grid gap-3 sm:grid-cols-2">
                  <label className="text-xs text-slate-400">OT<input required value={correctiveOtNumber} onChange={event => setCorrectiveOtNumber(event.target.value)} placeholder="Ej. 69001234" className="mt-2 h-10 w-full rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 text-sm text-white outline-none focus:border-sky-300/40" /></label>
                  <label className="text-xs text-slate-400">Equipo<input required value={correctiveEquipment} onChange={event => setCorrectiveEquipment(event.target.value)} placeholder="Ej. CA13 o CR02" className="mt-2 h-10 w-full rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 text-sm text-white outline-none focus:border-sky-300/40" /></label>
                </div>
                <label className="text-xs text-slate-400">Reserva <span className="text-slate-600">(opcional)</span><input value={correctiveReservation} onChange={event => setCorrectiveReservation(event.target.value)} placeholder="Código de reserva o dejar vacío" className="mt-2 h-10 w-full rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 text-sm text-white outline-none focus:border-sky-300/40" /></label>
                <label className="text-xs text-slate-400">Descripción de operación<textarea required value={correctiveDescription} onChange={event => setCorrectiveDescription(event.target.value)} placeholder="Describe el trabajo correctivo" rows={3} className="mt-2 w-full resize-none rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 py-2 text-sm text-white outline-none focus:border-sky-300/40" /></label>
              </> : <>
                <label className="text-xs text-slate-400">Pega una OT por línea<textarea autoFocus value={bulkCorrectiveText} onChange={event => { const value = event.target.value; setBulkCorrectiveText(value); setBulkCorrectivePreview(parseCorrectiveLines(value)); }} placeholder={'OT 6831949 R 3424500 CR CAMBIO DE FILTRO DE FRENO CR02\nOT 6831958 R 3424542 CR CAMBIO DE CLIP DE TUBERIA DE DIRECCION CF02'} rows={7} className="mt-2 w-full resize-y rounded-xl border border-sky-300/20 bg-white/[0.035] px-3 py-2 font-mono text-xs text-white outline-none focus:border-sky-300/40" /></label>
                <div className="rounded-xl border border-white/[0.07] bg-white/[0.025] p-3 text-xs">
                  <div className="font-semibold text-emerald-300">{bulkCorrectivePreview.rows.length} líneas listas para crear</div>
                  {bulkCorrectivePreview.invalid.length > 0 && <div className="mt-1 text-amber-300">{bulkCorrectivePreview.invalid.length} líneas no reconocidas; serán omitidas.</div>}
                  {bulkCorrectivePreview.rows.slice(0, 3).map(row => <div key={row.otNumber} className="mt-2 truncate text-slate-400"><span className="font-mono text-white">{row.otNumber}</span> · {row.reservation} · {row.equipmentCode} · {row.description}</div>)}
                </div>
              </>}
              <div className="flex justify-end gap-2 pt-1"><button type="button" onClick={() => setShowCorrectiveOt(false)} className="rounded-xl px-4 py-2.5 text-sm text-slate-400 hover:bg-white/5">Cancelar</button><button type="submit" disabled={createCorrectiveOtMutation.isPending || createCorrectiveOtsMutation.isPending} className="rounded-xl bg-sky-300 px-4 py-2.5 text-sm font-bold text-[#071014] disabled:opacity-50">{createCorrectiveOtMutation.isPending || createCorrectiveOtsMutation.isPending ? "Guardando…" : bulkCorrectiveMode ? `Crear ${bulkCorrectivePreview.rows.length} OT` : "Crear OT"}</button></div>
            </form>
          </div>
        </div>
      )}

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
                    <div className={`rounded-lg px-2.5 py-1.5 text-[11px] font-bold ${importPreview.detectedWeek ? "bg-emerald-400/10 text-emerald-300" : "bg-amber-300/10 text-amber-200"}`}>
                      {importPreview.detectedWeek
                        ? `Semana ${importPreview.detectedWeek.weekNumber} · ${importPreview.detectedWeek.year}`
                        : "Semana no detectada · revisar campos"}
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

      {editingCorrective && (
        <div className="fixed inset-0 z-[60] grid place-items-center bg-black/75 p-4 backdrop-blur-sm">
          <div className="w-full max-w-xl rounded-2xl border border-sky-300/20 bg-[#11191a] shadow-2xl">
            <div className="flex items-start justify-between border-b border-white/[0.07] p-5">
              <div>
                <div className="text-[10px] font-bold uppercase tracking-[0.18em] text-sky-300">OT correctiva manual</div>
                <h2 className="mt-1 font-display text-2xl text-white">Editar OT</h2>
                <p className="mt-1 text-sm text-slate-500">Actualiza los datos operativos de la OT.</p>
              </div>
              <button onClick={() => setEditingCorrective(null)} className="rounded-lg p-2 text-slate-500 hover:bg-white/5 hover:text-white" aria-label="Cerrar edición">
                <X className="h-5 w-5" />
              </button>
            </div>
            <form className="grid gap-4 p-5" onSubmit={event => {
              event.preventDefault();
              const { id, ...draft } = editingCorrective;
              updateCorrectiveOtMutation.mutate({ workOrderId: id, ...draft });
            }}>
              <div className="grid gap-3 sm:grid-cols-2">
                <label className="text-xs text-slate-400">OT
                  <input required value={editingCorrective.otNumber} onChange={event => setEditingCorrective(current => current ? { ...current, otNumber: event.target.value } : current)} className="mt-2 h-10 w-full rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 text-sm text-white outline-none focus:border-sky-300/40" />
                </label>
                <label className="text-xs text-slate-400">Equipo
                  <input required value={editingCorrective.equipmentCode} onChange={event => setEditingCorrective(current => current ? { ...current, equipmentCode: event.target.value } : current)} className="mt-2 h-10 w-full rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 text-sm text-white outline-none focus:border-sky-300/40" />
                </label>
              </div>
              <label className="text-xs text-slate-400">Reserva <span className="text-slate-600">(opcional)</span>
                <input value={editingCorrective.reservation} onChange={event => setEditingCorrective(current => current ? { ...current, reservation: event.target.value } : current)} className="mt-2 h-10 w-full rounded-xl border border-white/[0.08] bg-white/[0.035] px-3 text-sm text-white outline-none focus:border-sky-300/40" />
              </label>
              <label className="text-xs text-slate-400">Descripción de operación
                <textarea required rows={4} value={editingCorrective.description} onChange={event => setEditingCorrective(current => current ? { ...current, description: event.target.value } : current)} className="mt-2 w-full resize-none rounded-xl border border-white/[0.08] bg-white/[0.035] p-3 text-sm text-white outline-none focus:border-sky-300/40" />
              </label>
              <div className="flex justify-end gap-2 pt-1">
                <button type="button" onClick={() => setEditingCorrective(null)} className="rounded-xl px-4 py-2.5 text-sm text-slate-400 hover:bg-white/5">Cancelar</button>
                <button type="submit" disabled={updateCorrectiveOtMutation.isPending} className="rounded-xl bg-sky-300 px-4 py-2.5 text-sm font-bold text-[#071014] disabled:opacity-50">{updateCorrectiveOtMutation.isPending ? "Guardando…" : "Guardar cambios"}</button>
              </div>
            </form>
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
                    disabled
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
                            {item.actorName || "Usuario"} · Guardia {GUARD_LABEL[item.guard as Guard]}
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
                  {selectedWeek?.label} · Guardia {GUARD_LABEL[guard]}
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
                    Guardia {GUARD_LABEL[guard]}
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
                        Guardia {GUARD_LABEL[item]}
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
                  Crea el usuario, define su PIN y entrégale el acceso de su página.
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
              <form
                onSubmit={event => {
                  event.preventDefault();
                  if (newAccessPin.length < 8 || !newAccessLabel.trim()) {
                    toast.error("Indica un nombre y un PIN de al menos 8 caracteres.");
                    return;
                  }
                  createAccessCodeMutation.mutate({
                    label: newAccessLabel.trim(),
                    pin: newAccessPin,
                    role: newAccessRole,
                    guard: newAccessGuard,
                  });
                }}
                className="mb-5 rounded-xl border border-amber-300/15 bg-amber-300/[0.04] p-4"
              >
                <div className="mb-3 text-xs font-bold uppercase tracking-[0.15em] text-amber-200">
                  Crear usuario con PIN
                </div>
                <div className="grid gap-2 md:grid-cols-[1.1fr_1fr_0.9fr_auto]">
                  <input
                    value={newAccessLabel}
                    onChange={event => setNewAccessLabel(event.target.value)}
                    placeholder="Nombre del usuario"
                    className="h-10 rounded-lg border border-white/[0.08] bg-[#151e20] px-3 text-sm text-white outline-none focus:border-amber-300/40"
                  />
                  <div className="relative">
                    <input
                      value={newAccessPin}
                      onChange={event => setNewAccessPin(event.target.value)}
                      placeholder="PIN mínimo 8 caracteres"
                      type={showNewAccessPin ? "text" : "password"}
                      className="h-10 w-full rounded-lg border border-white/[0.08] bg-[#151e20] px-3 pr-10 text-sm text-white outline-none focus:border-amber-300/40"
                    />
                    <button
                      type="button"
                      aria-label={showNewAccessPin ? "Ocultar PIN" : "Mostrar PIN"}
                      title={showNewAccessPin ? "Ocultar PIN" : "Mostrar PIN"}
                      onClick={() => setShowNewAccessPin(current => !current)}
                      className="absolute right-1 top-1/2 -translate-y-1/2 rounded-md p-1.5 text-slate-500 hover:bg-white/5 hover:text-amber-200"
                    >
                      {showNewAccessPin ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                    </button>
                  </div>
                  <select
                    value={newAccessRole}
                    onChange={event => setNewAccessRole(event.target.value as "facilitator" | "supervisor")}
                    className="h-10 rounded-lg border border-white/[0.08] bg-[#151e20] px-3 text-sm text-white outline-none"
                  >
                    <option value="facilitator">Facilitador</option>
                    <option value="supervisor">Supervisor</option>
                  </select>
                  <select
                    value={newAccessGuard}
                    onChange={event => setNewAccessGuard(event.target.value as Guard)}
                    className="h-10 rounded-lg border border-white/[0.08] bg-[#151e20] px-3 text-sm text-white outline-none"
                  >
                    {GUARDS.map(item => <option key={item} value={item}>Guardia {GUARD_LABEL[item]}</option>)}
                  </select>
                  <button
                    type="submit"
                    disabled={createAccessCodeMutation.isPending}
                    className="h-10 rounded-lg bg-amber-300 px-4 text-xs font-bold text-[#0a0f10] disabled:opacity-40"
                  >
                    Crear acceso
                  </button>
                </div>
                <p className="mt-2 text-xs text-slate-500">
                  El PIN se guarda como hash; entrégaselo directamente al usuario y no se vuelve a mostrar en la lista.
                </p>
              </form>
              <div className="mb-5 overflow-x-auto rounded-xl border border-white/[0.07]">
                <table className="w-full min-w-[620px] text-left text-sm">
                  <thead className="bg-white/[0.035] text-[10px] uppercase tracking-[0.15em] text-slate-500">
                    <tr><th className="px-4 py-3">Usuario PIN</th><th className="px-4 py-3">Rol / guardia</th><th className="px-4 py-3">Estado</th><th className="px-4 py-3">Último acceso</th><th className="px-4 py-3" /></tr>
                  </thead>
                  <tbody className="divide-y divide-white/[0.05]">
                    {accessCodes.data?.map(code => (
                      <tr key={code.id}>
                        <td className="px-4 py-3 font-medium text-white">{code.label}</td>
                        <td className="px-4 py-3 text-slate-300">
                          <div>{code.role === "facilitator" ? "Facilitador" : "Supervisor"}</div>
                          <select
                            value={code.guard}
                            onChange={event => updateAccessCodeGuardMutation.mutate({ id: code.id, guard: event.target.value as Guard })}
                            disabled={updateAccessCodeGuardMutation.isPending}
                            className="mt-1 h-7 rounded-md border border-white/[0.08] bg-[#151e20] px-2 text-[11px] text-amber-200 outline-none"
                          >
                            {GUARDS.map(item => <option key={item} value={item}>Guardia {GUARD_LABEL[item]}</option>)}
                          </select>
                        </td>
                        <td className="px-4 py-3 text-xs">{code.active ? <span className="text-emerald-300">Activo</span> : <span className="text-slate-600">Desactivado</span>}</td>
                        <td className="px-4 py-3 text-xs text-slate-500">{formatDate(code.lastUsedAt)}</td>
                        <td className="px-4 py-3 text-right">{code.active && <button onClick={() => revokeAccessCodeMutation.mutate({ id: code.id })} className="rounded-lg border border-rose-400/20 px-2.5 py-1.5 text-[11px] font-semibold text-rose-300 hover:bg-rose-400/10">Desactivar</button>}</td>
                      </tr>
                    ))}
                    {!accessCodes.data?.length && <tr><td colSpan={5} className="px-4 py-6 text-center text-xs text-slate-500">Aún no hay usuarios con PIN creados.</td></tr>}
                  </tbody>
                </table>
              </div>
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
                        Guardia {GUARD_LABEL[to]}
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
