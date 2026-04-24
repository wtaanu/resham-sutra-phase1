import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import logoUrl from "./assets/resham-sutra-logo-small.png";

type Metric = {
  label: string;
  value: number;
};

type OperationsActionLinks = {
  interfaceUrl: string;
  enquiryFormUrl: string;
  lineItemsFormUrl: string;
  productsFormUrl: string;
};

type ProductDocument = {
  id: string;
  productId: string;
  productName: string;
  fileName: string;
  mimeType: string;
  uploadedAt: string;
  fileUrl: string;
};

type EnquiryRecord = {
  id: string;
  enquiryId: string;
  loggedDateTime: string;
  leadName: string;
  company: string;
  phone: string;
  email: string;
  address: string;
  state: string;
  city: string;
  pincode: string;
  destinationAddress: string;
  destinationState: string;
  destinationCity: string;
  destinationPincode: string;
  parserStatus: string;
  linkedCustomerId: string;
  quotations: string[];
  mappedProductDocuments: ProductDocument[];
  driveFolderUrl: string;
  requirementSummary: string;
  receiverWhatsappNumber: string;
};

type CustomerRecord = {
  id: string;
  clientId: string;
  customerName: string;
  company: string;
  phone: string;
  email: string;
  address: string;
  state: string;
  city: string;
  pincode: string;
  customerType: string;
  driveFolderUrl: string;
};

type QuotationRecord = {
  id: string;
  quotationNumber: string;
  loggedDateTime: string;
  linkedCustomerId: string;
  linkedEnquiryId: string;
  status: string;
  draftFileUrl: string;
  draftCreatedTime: string;
  finalPdfUrl: string;
  driveFolderUrl: string;
  preferredSendChannel: string;
  sentDate: string;
  whatsappSentDateTime: string;
  emailSentDateTime: string;
  finalPdfGeneratedAt?: string;
  lineItemCount?: number;
  sendQuotation?: boolean;
  sendReminder?: boolean;
  markAccepted?: boolean;
  markRejected?: boolean;
  reminderCount?: number;
  lastReminderDate?: string;
  nextReminderDate?: string;
};

type OrderRecord = {
  id: string;
  orderNumber: string;
  linkedCustomerId: string;
  linkedQuotationId: string;
  orderDate: string;
  orderValue: number;
  paymentStatus: string;
  deliveryStatus: string;
};

type ProductRecord = {
  id: string;
  productKey: string;
  model: string;
  name: string;
  narration: string;
  bulkSalePrice: number;
  mrp: number;
  gstPercent: number;
  transportCharge: number;
  sourceSheet: string;
  documents: ProductDocument[];
};

type OperationsResponse = {
  actions: OperationsActionLinks;
  metrics: Metric[];
  enquiries: EnquiryRecord[];
  customers: CustomerRecord[];
  quotations: QuotationRecord[];
  orders: OrderRecord[];
  products: ProductRecord[];
};

type ViewKey =
  | "dashboard"
  | "orders"
  | "sentQuotations"
  | "approvedQuotations"
  | "quotationDrafts"
  | "customers"
  | "enquiries"
  | "products";

type DashboardFeedItem = {
  id: string;
  title: string;
  subtitle: string;
  meta: string;
  href?: string;
};

type ActionState = {
  key: string;
  label: string;
  status: "idle" | "loading" | "error" | "success";
  message: string;
};

type AuthUser = {
  id: string;
  name: string;
  email: string;
};

type LoginFormState = {
  email: string;
  password: string;
};

type EntryMode = "enquiry" | "lineItems" | "productDocuments" | "enquiryDocuments" | null;

type EnquiryFormState = {
  linkedCustomerId: string;
  leadName: string;
  company: string;
  phone: string;
  email: string;
  receiverWhatsappNumber: string;
  address: string;
  state: string;
  city: string;
  pincode: string;
  destinationAddress: string;
  destinationState: string;
  destinationCity: string;
  destinationPincode: string;
  requirementSummary: string;
  requestedAsset: string;
  potentialProduct: string;
};

type EnquiryFieldErrors = Partial<Record<keyof EnquiryFormState, string>>;

type LineItemDraftRow = {
  id: string;
  productId: string;
  qty: string;
  totalAmount: string;
};

type PostalLookupResponse = Array<{
  Status: string;
  PostOffice?: Array<{
    District?: string;
    State?: string;
  }>;
}>;

type ChartPoint = {
  label: string;
  value: number;
};

type NavView = {
  key: ViewKey;
  label: string;
};

type PaginatedTableProps<T> = {
  eyebrow: string;
  title: string;
  subtitle?: string;
  rows: T[];
  columns: ReactNode;
  renderRow: (row: T) => ReactNode;
  emptyTitle: string;
  emptyBody: string;
  pageSize?: number;
};

const apiUrl = import.meta.env.VITE_API_URL ?? "http://localhost:4000";

const viewOptions: NavView[] = [
  { key: "dashboard", label: "Dashboard" },
  { key: "orders", label: "Orders" },
  { key: "sentQuotations", label: "Sent Quotations" },
  { key: "approvedQuotations", label: "Approved Quotations" },
  { key: "quotationDrafts", label: "Quotation Drafts" },
  { key: "customers", label: "Customers" },
  { key: "enquiries", label: "New Enquiries" },
  { key: "products", label: "Products" }
];

function formatCurrency(value: number) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0
  }).format(value);
}

function formatAmountInput(value: number) {
  if (!Number.isFinite(value)) {
    return "";
  }

  return value.toFixed(2).replace(/\.00$/, "");
}

function normalizePincodeInput(value: string) {
  return value.replace(/\D/g, "").slice(0, 6);
}

function normalizePhoneInput(value: string) {
  return value.replace(/\D/g, "").slice(0, 10);
}

function normalizeDecimalInput(value: string) {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }

  const normalized = trimmed.replace(/[^\d.]/g, "");
  const [whole = "", ...decimals] = normalized.split(".");
  const decimalPart = decimals.join("").slice(0, 2);
  return decimalPart ? `${whole}.${decimalPart}` : whole;
}

function isValidTenDigitPhone(value: string) {
  return normalizePhoneInput(value).length === 10;
}

function isValidEmail(value: string) {
  const trimmed = value.trim();
  return trimmed.includes("@");
}

function isValidPositiveAmount(value: string) {
  const trimmed = value.trim();
  if (!trimmed) {
    return true;
  }

  const parsed = Number(trimmed);
  return Number.isFinite(parsed) && parsed > 0;
}

function calculateLineItemAmounts(
  product: ProductRecord | undefined,
  qtyInput: string | number,
  totalAmountInput?: string
) {
  const qty = Math.max(1, Number(qtyInput || 0));
  const rate = Number(product?.bulkSalePrice || product?.mrp || 0);
  const transport = Number(product?.transportCharge || 0);
  const gstPercent = Number(product?.gstPercent || 0);
  const unitValue = Number((rate * qty).toFixed(2));
  const computedGstAmount = Number((((unitValue + transport) * gstPercent) / 100).toFixed(2));
  const computedTotalAmount = Number((unitValue + transport + computedGstAmount).toFixed(2));

  const normalizedInput = String(totalAmountInput ?? "").trim();
  const parsedOverride = Number(normalizedInput);
  const hasOverride = normalizedInput !== "" && Number.isFinite(parsedOverride) && parsedOverride > 0;
  const totalAmount = hasOverride ? Number(parsedOverride.toFixed(2)) : computedTotalAmount;
  const gstAmount = hasOverride
    ? Number((totalAmount - unitValue - transport).toFixed(2))
    : computedGstAmount;

  return {
    qty,
    rate,
    transport,
    gstPercent,
    unitValue,
    gstAmount,
    totalAmount,
    computedTotalAmount
  };
}

function formatDate(value: string) {
  if (!value) {
    return "Not set";
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return new Intl.DateTimeFormat("en-IN", {
    dateStyle: "medium"
  }).format(date);
}

function formatDateTime(value: string) {
  if (!value) {
    return "";
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return new Intl.DateTimeFormat("en-IN", {
    dateStyle: "medium",
    timeStyle: "short"
  }).format(date);
}

function statusTone(value: string) {
  const normalized = value.trim().toLowerCase();

  if (
    normalized.includes("approved") ||
    normalized.includes("sent") ||
    normalized.includes("ready") ||
    normalized.includes("processed")
  ) {
    return "good";
  }

  if (
    normalized.includes("draft") ||
    normalized.includes("review") ||
    normalized.includes("pending") ||
    normalized.includes("parsed")
  ) {
    return "warm";
  }

  if (normalized.includes("error") || normalized.includes("rejected") || normalized.includes("lost")) {
    return "danger";
  }

  return "neutral";
}

function ExternalAction({
  href,
  title,
  detail,
  disabled
}: {
  href: string;
  title: string;
  detail: string;
  disabled?: boolean;
}) {
  const className = disabled ? "action-card disabled" : "action-card";

  if (!href || disabled) {
    return (
      <article className={className}>
        <span className="action-kicker">Setup needed</span>
        <h3>{title}</h3>
        <p>{detail}</p>
      </article>
    );
  }

  return (
    <a className={className} href={href} target="_blank" rel="noreferrer">
      <span className="action-kicker">Open form</span>
      <h3>{title}</h3>
      <p>{detail}</p>
    </a>
  );
}

function DetailCard({
  title,
  rows
}: {
  title: string;
  rows: Array<{ label: string; value: string }>;
}) {
  return (
    <article className="detail-card">
      <h3>{title}</h3>
      <div className="detail-grid">
        {rows.map((row) => (
          <div key={`${title}-${row.label}`} className="detail-row">
            <span>{row.label}</span>
            <strong>{row.value || "-"}</strong>
          </div>
        ))}
      </div>
    </article>
  );
}

function createBlankEnquiryForm(): EnquiryFormState {
  return {
    linkedCustomerId: "",
    leadName: "",
    company: "",
    phone: "",
    email: "",
    receiverWhatsappNumber: "",
    address: "",
    state: "",
    city: "",
    pincode: "",
    destinationAddress: "",
    destinationState: "",
    destinationCity: "",
    destinationPincode: "",
    requirementSummary: "",
    requestedAsset: "",
    potentialProduct: ""
  };
}

function createLineItemRow(): LineItemDraftRow {
  return {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    productId: "",
    qty: "1",
    totalAmount: ""
  };
}

async function fetchPincodeDetails(pincode: string) {
  const trimmed = pincode.trim();
  if (!/^\d{6}$/.test(trimmed)) {
    return null;
  }

  const response = await fetch(`https://api.postalpincode.in/pincode/${trimmed}`);
  if (!response.ok) {
    throw new Error("Unable to reach pincode lookup service.");
  }

  const payload = (await response.json()) as PostalLookupResponse;
  const firstResult = payload[0];
  const firstOffice = firstResult?.PostOffice?.[0];

  if (firstResult?.Status !== "Success" || !firstOffice) {
    return null;
  }

  return {
    city: firstOffice.District || "",
    state: firstOffice.State || ""
  };
}

async function fileToBase64(file: File) {
  const buffer = await file.arrayBuffer();
  let binary = "";
  new Uint8Array(buffer).forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
}

function PaginatedTable<T>({
  eyebrow,
  title,
  subtitle,
  rows,
  columns,
  renderRow,
  emptyTitle,
  emptyBody,
  pageSize = 8
}: PaginatedTableProps<T>) {
  const [page, setPage] = useState(1);
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));

  useEffect(() => {
    setPage(1);
  }, [rows.length, title]);

  const start = (page - 1) * pageSize;
  const pageRows = rows.slice(start, start + pageSize);

  return (
    <section className="panel">
      <div className="panel-header panel-header-tight">
        <div>
          <p className="eyebrow">{eyebrow}</p>
          <h2>{title}</h2>
          {subtitle ? <p className="panel-subcopy">{subtitle}</p> : null}
        </div>
        <span className="table-meta">{rows.length} records</span>
      </div>

      {rows.length ? (
        <>
          <div className="table-shell">
            <table>
              <thead>
                <tr>{columns}</tr>
              </thead>
              <tbody>{pageRows.map(renderRow)}</tbody>
            </table>
          </div>
          <div className="pagination-bar">
            <span>
              Page {page} of {totalPages}
            </span>
            <div className="pagination-actions">
              <button type="button" disabled={page === 1} onClick={() => setPage(page - 1)}>
                Previous
              </button>
              <button
                type="button"
                disabled={page === totalPages}
                onClick={() => setPage(page + 1)}
              >
                Next
              </button>
            </div>
          </div>
        </>
      ) : (
        <article className="empty-card">
          <h3>{emptyTitle}</h3>
          <p>{emptyBody}</p>
        </article>
      )}
    </section>
  );
}

function MiniBarChart({ title, points }: { title: string; points: ChartPoint[] }) {
  const maxValue = Math.max(1, ...points.map((point) => point.value));

  return (
    <article className="chart-card">
      <div className="chart-header">
        <p className="eyebrow">Analytics</p>
        <h3>{title}</h3>
      </div>
      <div className="bar-chart">
        {points.map((point) => (
          <div className="bar-column" key={point.label}>
            <div className="bar-track">
              <div
                className="bar-fill"
                style={{ height: `${Math.max(10, (point.value / maxValue) * 100)}%` }}
              />
            </div>
            <strong>{point.value}</strong>
            <span>{point.label}</span>
          </div>
        ))}
      </div>
    </article>
  );
}

function DonutChart({
  title,
  value,
  total
}: {
  title: string;
  value: number;
  total: number;
}) {
  const safeTotal = Math.max(1, total);
  const angle = Math.min(360, (value / safeTotal) * 360);

  return (
    <article className="chart-card">
      <div className="chart-header">
        <p className="eyebrow">Snapshot</p>
        <h3>{title}</h3>
      </div>
      <div className="donut-wrap">
        <div
          className="donut-chart"
          style={{
            background: `conic-gradient(var(--accent) 0deg ${angle}deg, #e8edf5 ${angle}deg 360deg)`
          }}
        >
          <div className="donut-center">
            <strong>{value}</strong>
            <span>of {total}</span>
          </div>
        </div>
      </div>
    </article>
  );
}

export default function App() {
  const [currentUser, setCurrentUser] = useState<AuthUser | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authError, setAuthError] = useState("");
  const [loginForm, setLoginForm] = useState<LoginFormState>({
    email: "",
    password: ""
  });
  const [authActionState, setAuthActionState] = useState<ActionState | null>(null);
  const [profileOpen, setProfileOpen] = useState(false);
  const [operations, setOperations] = useState<OperationsResponse | null>(null);
  const [activeView, setActiveView] = useState<ViewKey>("dashboard");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [actionState, setActionState] = useState<ActionState | null>(null);
  const [entryMode, setEntryMode] = useState<EntryMode>(null);
  const [editingEnquiryId, setEditingEnquiryId] = useState("");
  const [enquiryForm, setEnquiryForm] = useState<EnquiryFormState>(createBlankEnquiryForm);
  const [enquiryFieldErrors, setEnquiryFieldErrors] = useState<EnquiryFieldErrors>({});
  const [destinationSameAsMain, setDestinationSameAsMain] = useState(false);
  const [customerSearchTerm, setCustomerSearchTerm] = useState("");
  const [customerDropdownOpen, setCustomerDropdownOpen] = useState(false);
  const [selectedQuotationId, setSelectedQuotationId] = useState("");
  const [lineItemRows, setLineItemRows] = useState<LineItemDraftRow[]>([createLineItemRow()]);
  const [selectedProductId, setSelectedProductId] = useState("");
  const [productDocFiles, setProductDocFiles] = useState<File[]>([]);
  const [selectedEnquiryId, setSelectedEnquiryId] = useState("");
  const [selectedDocumentIds, setSelectedDocumentIds] = useState<string[]>([]);
  const entryModalRef = useRef<HTMLElement | null>(null);
  const profileMenuRef = useRef<HTMLDivElement | null>(null);

  function dismissActionState() {
    setActionState(null);
  }

  function dismissAuthActionState() {
    setAuthActionState(null);
  }

  function renderBannerCloseButton(onClick: () => void) {
    return (
      <button type="button" className="banner-close" onClick={onClick}>
        Close
      </button>
    );
  }

  async function apiFetch(input: string, init?: RequestInit) {
    const response = await fetch(input, {
      ...init,
      credentials: "include",
      headers: {
        ...(init?.headers ?? {})
      }
    });

    if (response.status === 401) {
      setCurrentUser(null);
      setOperations(null);
      setProfileOpen(false);
      setActionState(null);
      setEntryMode(null);
      throw new Error("Your session expired. Please sign in again.");
    }

    return response;
  }

  async function loadSession() {
    try {
      setAuthLoading(true);
      const response = await fetch(`${apiUrl}/api/auth/session`, {
        credentials: "include"
      });

      if (!response.ok) {
        setCurrentUser(null);
        return;
      }

      const payload = (await response.json()) as { user: AuthUser };
      setCurrentUser(payload.user);
    } catch {
      setCurrentUser(null);
    } finally {
      setAuthLoading(false);
    }
  }

  async function handleLogin() {
    try {
      setAuthError("");
      setAuthActionState({
        key: "auth-login",
        label: "Sign In",
        status: "loading",
        message: "Signing you into ReshamSutra..."
      });

      const response = await fetch(`${apiUrl}/api/auth/login`, {
        method: "POST",
        credentials: "include",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify(loginForm)
      });

      const payload = (await response.json()) as { message?: string; user?: AuthUser };
      if (!response.ok || !payload.user) {
        throw new Error(payload.message || "Invalid email or password.");
      }

      setCurrentUser(payload.user);
      setLoginForm({
        email: payload.user.email,
        password: ""
      });
      setAuthActionState({
        key: "auth-login",
        label: "Sign In",
        status: "success",
        message: "Signed in successfully."
      });
    } catch (loginError) {
      const message = loginError instanceof Error ? loginError.message : "Failed to sign in";
      setAuthError(message);
      setAuthActionState({
        key: "auth-login",
        label: "Sign In",
        status: "error",
        message
      });
    }
  }

  async function handleLogout() {
    try {
      setAuthActionState({
        key: "auth-logout",
        label: "Logout",
        status: "loading",
        message: "Signing you out..."
      });

      await fetch(`${apiUrl}/api/auth/logout`, {
        method: "POST",
        credentials: "include"
      });
    } finally {
      setCurrentUser(null);
      setOperations(null);
      setProfileOpen(false);
      setActionState(null);
      setEntryMode(null);
      setError("");
      setLoading(false);
      setAuthActionState({
        key: "auth-logout",
        label: "Logout",
        status: "success",
        message: "You have been signed out."
      });
    }
  }

  useEffect(() => {
    void loadSession();
  }, []);

  useEffect(() => {
    function handlePointerDown(event: MouseEvent) {
      if (!profileMenuRef.current?.contains(event.target as Node)) {
        setProfileOpen(false);
      }
    }

    document.addEventListener("mousedown", handlePointerDown);
    return () => {
      document.removeEventListener("mousedown", handlePointerDown);
    };
  }, []);

  useEffect(() => {
    if (!actionState || actionState.status === "loading") {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      setActionState((current) => (current === actionState ? null : current));
    }, 5000);

    return () => window.clearTimeout(timeoutId);
  }, [actionState]);

  useEffect(() => {
    if (!authActionState || authActionState.status === "loading") {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      setAuthActionState((current) => (current === authActionState ? null : current));
    }, 5000);

    return () => window.clearTimeout(timeoutId);
  }, [authActionState]);

  useEffect(() => {
    if (!error || !operations) {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      setError("");
    }, 5000);

    return () => window.clearTimeout(timeoutId);
  }, [error, operations]);

  useEffect(() => {
    if (authLoading || !currentUser) {
      setLoading(false);
      return;
    }

    let cancelled = false;

    async function loadOperations(isBackgroundRefresh = false) {
      try {
        if (!isBackgroundRefresh) {
          setLoading(true);
        }
        setError("");

        const response = await apiFetch(`${apiUrl}/api/operations`);
        if (!response.ok) {
          throw new Error(`Operations API returned ${response.status}`);
        }

        const data = (await response.json()) as OperationsResponse;
        if (!cancelled) {
          setOperations(data);
        }
      } catch (loadError) {
        if (!cancelled) {
          const message =
            loadError instanceof Error ? loadError.message : "Failed to load dashboard";
          setError(message);
          if (message.includes("sign in again")) {
            setAuthError(message);
          }
        }
      } finally {
        if (!cancelled && !isBackgroundRefresh) {
          setLoading(false);
        }
      }
    }

    void loadOperations();

    const intervalId = window.setInterval(() => {
      void loadOperations(true);
    }, 15000);

    const handleFocus = () => {
      void loadOperations(true);
    };

    window.addEventListener("focus", handleFocus);

    return () => {
      cancelled = true;
      window.clearInterval(intervalId);
      window.removeEventListener("focus", handleFocus);
    };
  }, [authLoading, currentUser]);

  useEffect(() => {
    if (!destinationSameAsMain) {
      return;
    }

    setEnquiryForm((current) => ({
      ...current,
      destinationAddress: current.address,
      destinationPincode: current.pincode,
      destinationState: current.state,
      destinationCity: current.city
    }));
  }, [
    destinationSameAsMain,
    enquiryForm.address,
    enquiryForm.pincode,
    enquiryForm.state,
    enquiryForm.city
  ]);

  async function refreshOperations(isBackgroundRefresh = true) {
    if (!currentUser) {
      return;
    }

    try {
      if (!isBackgroundRefresh) {
        setLoading(true);
      }
      const response = await apiFetch(`${apiUrl}/api/operations`);
      if (!response.ok) {
        throw new Error(`Operations API returned ${response.status}`);
      }
      const data = (await response.json()) as OperationsResponse;
      setOperations(data);
      setError("");
    } catch (loadError) {
      const message =
        loadError instanceof Error ? loadError.message : "Failed to refresh dashboard";
      if (!operations) {
        setError(message);
      }
      if (message.includes("sign in again")) {
        setAuthError(message);
      }
    } finally {
      if (!isBackgroundRefresh) {
        setLoading(false);
      }
    }
  }

  async function handleCreateCustomer(enquiryId: string) {
    try {
      setActionState({
        key: enquiryId,
        label: "Create Customer",
        status: "loading",
        message: "Creating customer and quotation shell..."
      });

      const response = await apiFetch(`${apiUrl}/api/actions/enquiries/${enquiryId}/create-customer`, {
        method: "POST"
      });

      if (!response.ok) {
        const payload = (await response.json()) as { message?: string };
        throw new Error(payload.message || "Failed to create customer");
      }

      await refreshOperations(false);
      setActionState({
        key: enquiryId,
        label: "Create Customer",
        status: "success",
        message: "Customer linked successfully."
      });
    } catch (actionError) {
      const message =
        actionError instanceof Error ? actionError.message : "Failed to create customer";
      setActionState({
        key: enquiryId,
        label: "Create Customer",
        status: "error",
        message
      });
    }
  }

  async function handleGenerateFinalPdf(quotationId: string) {
    try {
      setActionState({
        key: `pdf-generate-${quotationId}`,
        label: "Generate Final PDF",
        status: "loading",
        message: "Creating final quotation PDF..."
      });

      const response = await apiFetch(`${apiUrl}/api/actions/quotations/${quotationId}/generate-final-pdf`, {
        method: "POST"
      });

      if (!response.ok) {
        const payload = (await response.json()) as { message?: string };
        throw new Error(payload.message || "Failed to generate final PDF");
      }

      await refreshOperations(false);
      setActiveView("approvedQuotations");
      setActionState({
        key: `pdf-generate-${quotationId}`,
        label: "Generate Final PDF",
        status: "success",
        message: "Final PDF generated successfully."
      });
    } catch (actionError) {
      const message =
        actionError instanceof Error ? actionError.message : "Failed to generate final PDF";
      setActionState({
        key: `pdf-generate-${quotationId}`,
        label: "Generate Final PDF",
        status: "error",
        message
      });
    }
  }

  async function handleSendQuotation(quotationId: string, channel: "email" | "whatsapp") {
    const actionKey = `quotation-send-${channel}-${quotationId}`;
    const label = channel === "email" ? "Send on Email" : "Send on WhatsApp";

    try {
      setActionState({
        key: actionKey,
        label,
        status: "loading",
        message:
          channel === "email"
            ? "Sending quotation on email..."
            : "Sending quotation on WhatsApp..."
      });

      const response = await apiFetch(`${apiUrl}/api/actions/quotations/${quotationId}/send-${channel}`, {
        method: "POST"
      });

      if (!response.ok) {
        const payload = (await response.json()) as { message?: string };
        throw new Error(payload.message || `Failed to ${label.toLowerCase()}`);
      }

      await refreshOperations(false);
      setActiveView("sentQuotations");
      setActionState({
        key: actionKey,
        label,
        status: "success",
        message:
          channel === "email"
            ? "Quotation emailed successfully."
            : "Quotation sent on WhatsApp successfully."
      });
    } catch (actionError) {
      const message =
        actionError instanceof Error ? actionError.message : `Failed to ${label.toLowerCase()}`;
      setActionState({
        key: actionKey,
        label,
        status: "error",
        message
      });
    }
  }

  async function handleRegenerateQuotationDraft(quotationId: string) {
    const actionKey = `quotation-regenerate-${quotationId}`;

    try {
      setActionState({
        key: actionKey,
        label: "Regenerate Draft",
        status: "loading",
        message: "Refreshing quotation draft and syncing files..."
      });

      const response = await apiFetch(`${apiUrl}/api/actions/quotations/${quotationId}/regenerate-draft`, {
        method: "POST"
      });

      if (!response.ok) {
        const payload = (await response.json()) as { message?: string };
        throw new Error(payload.message || "Failed to regenerate quotation draft");
      }

      await refreshOperations(false);
      setActiveView("quotationDrafts");
      setActionState({
        key: actionKey,
        label: "Regenerate Draft",
        status: "success",
        message: "Quotation draft regenerated successfully."
      });
    } catch (actionError) {
      const message =
        actionError instanceof Error ? actionError.message : "Failed to regenerate quotation draft";
      setActionState({
        key: actionKey,
        label: "Regenerate Draft",
        status: "error",
        message
      });
    }
  }

  function handleLinkAction(actionKey: string, label: string) {
    setActionState({
      key: actionKey,
      label,
      status: "loading",
      message: `${label} is opening...`
    });

    window.setTimeout(() => {
      setActionState((current) =>
        current?.key === actionKey
          ? {
              key: actionKey,
              label,
              status: "success",
              message: `${label} opened in a new tab.`
            }
          : current
      );
    }, 1200);
  }

  function openEnquiryEntry() {
    setEntryMode("enquiry");
    setEditingEnquiryId("");
    setEnquiryForm(createBlankEnquiryForm());
    setEnquiryFieldErrors({});
    setDestinationSameAsMain(false);
    setCustomerSearchTerm("");
    setCustomerDropdownOpen(false);
  }

  function openEnquiryEdit(enquiry: EnquiryRecord) {
    const customer = customerLookup.get(enquiry.linkedCustomerId);
    setEntryMode("enquiry");
    setEditingEnquiryId(enquiry.id);
    setEnquiryForm({
      linkedCustomerId: enquiry.linkedCustomerId,
      leadName: enquiry.leadName,
      company: enquiry.company,
      phone: enquiry.phone,
      email: enquiry.email,
      receiverWhatsappNumber: enquiry.receiverWhatsappNumber,
      address: enquiry.address,
      state: enquiry.state,
      city: enquiry.city,
      pincode: enquiry.pincode,
      destinationAddress: enquiry.destinationAddress,
      destinationState: enquiry.destinationState,
      destinationCity: enquiry.destinationCity,
      destinationPincode: enquiry.destinationPincode,
      requirementSummary: enquiry.requirementSummary,
      requestedAsset: "",
      potentialProduct: ""
    });
    setEnquiryFieldErrors({});
    setDestinationSameAsMain(
      Boolean(
        enquiry.address &&
          enquiry.address === enquiry.destinationAddress &&
          enquiry.state === enquiry.destinationState &&
          enquiry.city === enquiry.destinationCity &&
          enquiry.pincode === enquiry.destinationPincode
      )
    );
    setCustomerSearchTerm(
      customer ? `${customer.clientId || "No ID"} - ${customer.customerName || "Unnamed customer"}` : ""
    );
    setCustomerDropdownOpen(false);
  }

  function openLineItemEntry(quotationId = "") {
    setEntryMode("lineItems");
    setSelectedQuotationId(quotationId);
    setLineItemRows([createLineItemRow()]);
  }

  function openProductDocuments(productId = "") {
    setEntryMode("productDocuments");
    setSelectedProductId(productId);
    setProductDocFiles([]);
  }

  function openEnquiryDocuments(enquiryId: string) {
    setEntryMode("enquiryDocuments");
    setSelectedEnquiryId(enquiryId);
    setSelectedDocumentIds([]);
  }

  function closeEntryPanel() {
    setEntryMode(null);
    setEditingEnquiryId("");
    setEnquiryFieldErrors({});
    setDestinationSameAsMain(false);
    setCustomerSearchTerm("");
    setCustomerDropdownOpen(false);
  }

  const customerLookup = useMemo(() => {
    return new Map((operations?.customers ?? []).map((customer) => [customer.id, customer]));
  }, [operations?.customers]);

  const sortedCustomers = useMemo(() => {
    return [...(operations?.customers ?? [])].sort((left, right) => {
      const leftName = (left.customerName || "").toLowerCase();
      const rightName = (right.customerName || "").toLowerCase();
      return leftName.localeCompare(rightName) || (left.clientId || "").localeCompare(right.clientId || "");
    });
  }, [operations?.customers]);

  const enquiryLookup = useMemo(() => {
    return new Map((operations?.enquiries ?? []).map((enquiry) => [enquiry.id, enquiry]));
  }, [operations?.enquiries]);

  const quotationLookup = useMemo(() => {
    return new Map((operations?.quotations ?? []).map((quotation) => [quotation.id, quotation]));
  }, [operations?.quotations]);

  const availableQuotationOptions = useMemo(() => {
    return (operations?.quotations ?? []).filter((quotation) =>
      ["Parsed", "Ready for Draft", "Draft", "Under Review"].includes(quotation.status || "")
    );
  }, [operations?.quotations]);

  const filteredCustomers = useMemo(() => {
    const term = customerSearchTerm.trim().toLowerCase();
    if (!term) {
      return sortedCustomers;
    }

    return sortedCustomers.filter((customer) =>
      `${customer.clientId} ${customer.customerName} ${customer.company} ${customer.phone} ${customer.email}`
        .toLowerCase()
        .includes(term)
    );
  }, [customerSearchTerm, sortedCustomers]);

  function clearEnquiryFieldError(field: keyof EnquiryFormState) {
    setEnquiryFieldErrors((current) => {
      if (!current[field]) {
        return current;
      }

      const next = { ...current };
      delete next[field];
      return next;
    });
  }

  function handleCustomerAutofill(customerId: string) {
    const customer = customerLookup.get(customerId);
    if (customer) {
      setCustomerSearchTerm(`${customer.clientId || "No ID"} - ${customer.customerName || "Unnamed customer"}`);
    }
    setCustomerDropdownOpen(false);

    setEnquiryForm((current) => ({
      ...current,
      linkedCustomerId: customerId,
      leadName: customer?.customerName || current.leadName,
      company: customer?.company || current.company,
      phone: customer?.phone || current.phone,
      email: customer?.email || current.email,
      address: customer?.address || current.address,
      state: customer?.state || current.state,
      city: customer?.city || current.city,
      pincode: customer?.pincode || current.pincode
    }));
  }

  async function handleSubmitEnquiry() {
    const isEditing = Boolean(editingEnquiryId);
    const formActionLabel = isEditing ? "Update Enquiry" : "Create Enquiry";
    const normalizedPhone = normalizePhoneInput(enquiryForm.phone);
    const normalizedReceiverWhatsapp = normalizePhoneInput(enquiryForm.receiverWhatsappNumber);
    const normalizedEmail = enquiryForm.email.trim();
    const nextFieldErrors: EnquiryFieldErrors = {};

    if (enquiryForm.phone.trim() && !isValidTenDigitPhone(enquiryForm.phone)) {
      nextFieldErrors.phone = "Phone number must be exactly 10 digits.";
      setEnquiryFieldErrors(nextFieldErrors);
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "error",
        message: nextFieldErrors.phone
      });
      return;
    }

    if (normalizedEmail && !isValidEmail(normalizedEmail)) {
      nextFieldErrors.email = "Email must include @.";
      setEnquiryFieldErrors(nextFieldErrors);
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "error",
        message: nextFieldErrors.email
      });
      return;
    }

    if (enquiryForm.receiverWhatsappNumber.trim() && !isValidTenDigitPhone(enquiryForm.receiverWhatsappNumber)) {
      nextFieldErrors.receiverWhatsappNumber = "Receiver WhatsApp number must be exactly 10 digits.";
      setEnquiryFieldErrors(nextFieldErrors);
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "error",
        message: nextFieldErrors.receiverWhatsappNumber
      });
      return;
    }

    if (!/^\d{6}$/.test(normalizePincodeInput(enquiryForm.pincode))) {
      nextFieldErrors.pincode = "Main pincode must be a valid 6-digit number.";
      setEnquiryFieldErrors(nextFieldErrors);
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "error",
        message: nextFieldErrors.pincode
      });
      return;
    }

    if (!/^\d{6}$/.test(normalizePincodeInput(enquiryForm.destinationPincode))) {
      nextFieldErrors.destinationPincode = "Destination pincode must be a valid 6-digit number.";
      setEnquiryFieldErrors(nextFieldErrors);
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "error",
        message: nextFieldErrors.destinationPincode
      });
      return;
    }

    setEnquiryFieldErrors({});

    if (
      !enquiryForm.address.trim() ||
      !enquiryForm.pincode.trim() ||
      !enquiryForm.state.trim() ||
      !enquiryForm.city.trim()
    ) {
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "error",
        message: "Address, pincode, state, and city are required."
      });
      return;
    }

    if (
      !enquiryForm.destinationAddress.trim() ||
      !enquiryForm.destinationPincode.trim() ||
      !enquiryForm.destinationState.trim() ||
      !enquiryForm.destinationCity.trim()
    ) {
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "error",
        message: "Destination address, pincode, state, and city are required."
      });
      return;
    }

    try {
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "loading",
        message: isEditing
          ? "Updating enquiry and refreshing customer, folder, and quotation flow..."
          : "Saving enquiry to Airtable..."
      });

      const response = await apiFetch(
        isEditing
          ? `${apiUrl}/api/portal/enquiries/${editingEnquiryId}`
          : `${apiUrl}/api/portal/enquiries`,
        {
          method: isEditing ? "PATCH" : "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            linkedCustomerId: enquiryForm.linkedCustomerId,
            leadName: enquiryForm.leadName,
            company: enquiryForm.company,
            phone: normalizedPhone,
            email: normalizedEmail,
            receiverWhatsappNumber: normalizedReceiverWhatsapp,
            address: enquiryForm.address,
            state: enquiryForm.state,
            city: enquiryForm.city,
            pincode: normalizePincodeInput(enquiryForm.pincode),
            destinationAddress: enquiryForm.destinationAddress,
            destinationState: enquiryForm.destinationState,
            destinationCity: enquiryForm.destinationCity,
            destinationPincode: normalizePincodeInput(enquiryForm.destinationPincode),
            requirementSummary: enquiryForm.requirementSummary,
            requestedAsset: enquiryForm.requestedAsset,
            potentialProduct: enquiryForm.potentialProduct
          })
        }
      );

      const payload = (await response.json()) as { message?: string };

      if (!response.ok) {
        throw new Error(payload.message || (isEditing ? "Failed to update enquiry" : "Failed to create enquiry"));
      }

      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "success",
        message:
          payload.message ||
          (isEditing
            ? "Enquiry updated and provisioning flow refreshed successfully."
            : "Enquiry created successfully.")
      });
      setEnquiryForm(createBlankEnquiryForm());
      closeEntryPanel();
      setActiveView("enquiries");
      void refreshOperations(true);
    } catch (submitError) {
      const message =
        submitError instanceof Error
          ? submitError.message
          : isEditing
            ? "Failed to update enquiry"
            : "Failed to create enquiry";
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "error",
        message
      });
    }
  }

  function updateLineItemRow(rowId: string, field: "productId" | "qty" | "totalAmount", value: string) {
    setLineItemRows((current) =>
      current.map((row) => {
        if (row.id !== rowId) {
          return row;
        }

        const nextValue = field === "totalAmount" ? normalizeDecimalInput(value) : value;
        const nextRow = { ...row, [field]: nextValue };

        if (field === "totalAmount") {
          return nextRow;
        }

        const product = operations?.products.find((item) => item.id === nextRow.productId);
        const amounts = calculateLineItemAmounts(product, nextRow.qty);

        return {
          ...nextRow,
          totalAmount: nextRow.productId ? formatAmountInput(amounts.totalAmount) : ""
        };
      })
    );
  }

  async function autofillFromPincode(
    pincode: string,
    type: "source" | "destination"
  ) {
    try {
      const details = await fetchPincodeDetails(pincode);
      if (!details) {
        return;
      }

      setEnquiryForm((current) =>
        type === "source"
          ? {
              ...current,
              pincode,
              city: current.city || details.city,
              state: current.state || details.state
            }
          : {
              ...current,
              destinationPincode: pincode,
              destinationCity: current.destinationCity || details.city,
              destinationState: current.destinationState || details.state
            }
      );
    } catch (lookupError) {
      setActionState({
        key: `${type}-pincode`,
        label: "Pincode lookup",
        status: "error",
        message:
          lookupError instanceof Error
            ? lookupError.message
            : "Unable to auto-fill city and state from pincode."
      });
    }
  }

  function addLineItemRow() {
    setLineItemRows((current) => [...current, createLineItemRow()]);
  }

  function removeLineItemRow(rowId: string) {
    setLineItemRows((current) => (current.length > 1 ? current.filter((row) => row.id !== rowId) : current));
  }

  async function handleSubmitLineItems() {
    if (lineItemRows.some((row) => !row.productId)) {
      setActionState({
        key: "portal-line-items",
        label: "Create Line Items",
        status: "error",
        message: "Select a product for every line item before creating them."
      });
      return;
    }

    if (
      lineItemRows.some(
        (row) => row.totalAmount.trim() !== "" && !isValidPositiveAmount(row.totalAmount)
      )
    ) {
      setActionState({
        key: "portal-line-items",
        label: "Create Line Items",
        status: "error",
        message: "Total amount must be a number greater than zero. Decimals are allowed."
      });
      return;
    }

    try {
      setActionState({
        key: "portal-line-items",
        label: "Create Line Items",
        status: "loading",
        message: "Creating quotation line items..."
      });

      const response = await apiFetch(`${apiUrl}/api/portal/quotation-line-items`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          quotationId: selectedQuotationId,
          items: lineItemRows.map((row) => ({
            productId: row.productId,
            qty: row.qty,
            totalAmount: row.totalAmount
          }))
        })
      });

      const payload = (await response.json()) as { message?: string; quotationNumber?: string; createdCount?: number };

      if (!response.ok) {
        throw new Error(payload.message || "Failed to create line items");
      }

      setActionState({
        key: "portal-line-items",
        label: "Create Line Items",
        status: "success",
        message: `${payload.createdCount || lineItemRows.length} line items created for ${payload.quotationNumber || "quotation"}.`
      });
      setLineItemRows([createLineItemRow()]);
      closeEntryPanel();
      setActiveView("quotationDrafts");
      void refreshOperations(true);
    } catch (submitError) {
      const message =
        submitError instanceof Error ? submitError.message : "Failed to create line items";
      setActionState({
        key: "portal-line-items",
        label: "Create Line Items",
        status: "error",
        message
      });
    }
  }

  async function handleUploadProductDocuments() {
    if (!selectedProductId || !productDocFiles.length) {
      setActionState({
        key: "product-documents",
        label: "Upload Documents",
        status: "error",
        message: "Choose a product and at least one document first."
      });
      return;
    }

    try {
      setActionState({
        key: "product-documents",
        label: "Upload Documents",
        status: "loading",
        message: "Uploading product documents..."
      });

      const files = await Promise.all(
        productDocFiles.map(async (file) => ({
          name: file.name,
          mimeType: file.type || "application/octet-stream",
          contentBase64: await fileToBase64(file)
        }))
      );

      const response = await apiFetch(`${apiUrl}/api/portal/products/${selectedProductId}/documents`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ files })
      });

      const payload = (await response.json()) as { message?: string; uploadedCount?: number };
      if (!response.ok) {
        throw new Error(payload.message || "Failed to upload product documents");
      }

      await refreshOperations(false);
      setProductDocFiles([]);
      setActionState({
        key: "product-documents",
        label: "Upload Documents",
        status: "success",
        message: `${payload.uploadedCount || files.length} document(s) uploaded successfully.`
      });
    } catch (uploadError) {
      const message =
        uploadError instanceof Error ? uploadError.message : "Failed to upload product documents";
      setActionState({
        key: "product-documents",
        label: "Upload Documents",
        status: "error",
        message
      });
    }
  }

  async function handleSendProductDocuments() {
    if (!selectedEnquiryId || !selectedDocumentIds.length) {
      setActionState({
        key: "send-product-documents",
        label: "Send Product Documents",
        status: "error",
        message: "Select at least one document to prepare the send action."
      });
      return;
    }

    try {
      setActionState({
        key: "send-product-documents",
        label: "Send Product Documents",
        status: "loading",
        message: "Preparing product documents for email and WhatsApp..."
      });

      const response = await apiFetch(`${apiUrl}/api/actions/enquiries/${selectedEnquiryId}/send-product-documents`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          documentIds: selectedDocumentIds
        })
      });

      const payload = (await response.json()) as { message?: string; customer?: { email?: string; phone?: string } };
      if (!response.ok) {
        throw new Error(payload.message || "Failed to prepare product document send");
      }

      setActionState({
        key: "send-product-documents",
        label: "Send Product Documents",
        status: "success",
        message:
          payload.message ||
          `Prepared ${selectedDocumentIds.length} document(s) for ${payload.customer?.email || payload.customer?.phone || "the customer"}.`
      });
    } catch (sendError) {
      const message =
        sendError instanceof Error ? sendError.message : "Failed to prepare product document send";
      setActionState({
        key: "send-product-documents",
        label: "Send Product Documents",
        status: "error",
        message
      });
    }
  }

  const dashboardFeed = useMemo<DashboardFeedItem[]>(() => {
    if (!operations) {
      return [];
    }

    return [
      ...operations.enquiries.slice(0, 2).map((enquiry) => ({
        id: enquiry.id,
        title: enquiry.leadName || enquiry.enquiryId,
        subtitle: enquiry.company || enquiry.requirementSummary || "New enquiry",
        meta: enquiry.parserStatus || "New",
        href: enquiry.driveFolderUrl || undefined
      })),
      ...operations.quotations.slice(0, 2).map((quotation) => ({
        id: quotation.id,
        title: quotation.quotationNumber,
        subtitle: customerLookup.get(quotation.linkedCustomerId)?.customerName || "Quotation workflow",
        meta: quotation.status || "Draft",
        href: quotation.draftFileUrl || quotation.finalPdfUrl || quotation.driveFolderUrl || undefined
      }))
    ];
  }, [customerLookup, operations]);

  const selectedQuotation = selectedQuotationId ? quotationLookup.get(selectedQuotationId) : undefined;
  const selectedProduct = selectedProductId
    ? operations?.products.find((product) => product.id === selectedProductId)
    : undefined;
  const selectedEnquiry = selectedEnquiryId ? enquiryLookup.get(selectedEnquiryId) : undefined;
  const isSavingEnquiry = actionState?.key === "portal-enquiry" && actionState.status === "loading";
  const isCreatingLineItems =
    actionState?.key === "portal-line-items" && actionState.status === "loading";
  const isUploadingProductDocuments =
    actionState?.key === "product-documents" && actionState.status === "loading";
  const isSendingProductDocuments =
    actionState?.key === "send-product-documents" && actionState.status === "loading";
  const hasInvalidLineItems = lineItemRows.some(
    (row) =>
      !row.productId ||
      Number(row.qty || 0) <= 0 ||
      (row.totalAmount.trim() !== "" && !isValidPositiveAmount(row.totalAmount))
  );
  const popupActionState = useMemo(() => {
    if (!actionState || !entryMode) {
      return null;
    }

    switch (entryMode) {
      case "enquiry":
        return actionState.key === "portal-enquiry" || actionState.key.endsWith("-pincode")
          ? actionState
          : null;
      case "productDocuments":
        return actionState.key === "product-documents" ? actionState : null;
      case "enquiryDocuments":
        return actionState.key === "send-product-documents" ? actionState : null;
      case "lineItems":
        return actionState.key === "portal-line-items" ? actionState : null;
      default:
        return null;
    }
  }, [actionState, entryMode]);

  useEffect(() => {
    if (!popupActionState || popupActionState.status !== "error") {
      return;
    }

    entryModalRef.current?.scrollTo({ top: 0, behavior: "smooth" });
  }, [popupActionState]);

  const renderEntryPanel = () => {
    if (!operations || !entryMode) {
      return null;
    }

    if (entryMode === "enquiry") {
      return (
        <section className="entry-overlay">
          <div className="entry-backdrop" onClick={closeEntryPanel} />
          <section className="panel entry-panel entry-modal" ref={entryModalRef}>
            <div className="panel-header panel-header-tight">
              <div>
                <p className="eyebrow">Portal Entry</p>
                <h2>{editingEnquiryId ? "Edit enquiry" : "Create enquiry"}</h2>
                <p className="panel-subcopy">
                  {editingEnquiryId
                    ? "Update the enquiry and re-run customer, folder, and quotation provisioning in one save."
                    : "Capture the enquiry here and push it straight into Airtable."}
                </p>
              </div>
              <button className="entry-close" type="button" onClick={closeEntryPanel}>
                Close
              </button>
            </div>

            {popupActionState ? (
              <section className={`action-banner entry-action-banner ${popupActionState.status}`}>
                <strong>{popupActionState.label}</strong>
                <span>{popupActionState.message}</span>
                {renderBannerCloseButton(dismissActionState)}
              </section>
            ) : null}

            <div className="form-grid">
              <label className="form-span-2">
                <span>Existing customer</span>
                <div className="searchable-dropdown">
                  <div className="searchable-dropdown-input">
                    <input
                      value={customerSearchTerm}
                      placeholder="Search by Client ID, name, phone, or email"
                      onFocus={() => setCustomerDropdownOpen(true)}
                      onClick={() => setCustomerDropdownOpen(true)}
                      onBlur={() => {
                        window.setTimeout(() => setCustomerDropdownOpen(false), 120);
                      }}
                      onChange={(event) => {
                        const value = event.target.value;
                        setCustomerSearchTerm(value);
                        setCustomerDropdownOpen(true);
                        if (!value.trim()) {
                          setEnquiryForm((current) => ({ ...current, linkedCustomerId: "" }));
                        }
                      }}
                    />
                    <button
                      type="button"
                      className="search-dropdown-toggle"
                      aria-label="Browse existing customers"
                      onMouseDown={(event) => {
                        event.preventDefault();
                        setCustomerDropdownOpen((current) => !current);
                      }}
                    >
                      {customerDropdownOpen ? "▲" : "▼"}
                    </button>
                  </div>
                  <span className="search-dropdown-hint">
                    Search or click the arrow to browse existing customers.
                  </span>
                  {customerDropdownOpen ? (
                    <div className="search-dropdown-list">
                      {filteredCustomers.length ? filteredCustomers.slice(0, 8).map((customer) => (
                        <button
                          key={customer.id}
                          className="search-dropdown-option"
                          type="button"
                          onMouseDown={(event) => {
                            event.preventDefault();
                            handleCustomerAutofill(customer.id);
                          }}
                        >
                          <strong>{`${customer.clientId || "No ID"} - ${customer.customerName || "Unnamed customer"}`}</strong>
                          <span>
                            {[customer.phone, customer.email, customer.company].filter(Boolean).join(" • ") || "No extra details"}
                          </span>
                        </button>
                      )) : <div className="search-dropdown-empty">No matching customers found.</div>}
                    </div>
                  ) : null}
                </div>
              </label>
            <label>
              <span>Lead name *</span>
              <input
                value={enquiryForm.leadName}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, leadName: event.target.value }))
                }
              />
            </label>
            <label>
              <span>Company</span>
              <input
                value={enquiryForm.company}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, company: event.target.value }))
                }
              />
            </label>
            <label>
              <span>Phone</span>
              <input
                inputMode="numeric"
                maxLength={10}
                value={enquiryForm.phone}
                onChange={(event) => {
                  clearEnquiryFieldError("phone");
                  setEnquiryForm((current) => ({
                    ...current,
                    phone: normalizePhoneInput(event.target.value)
                  }));
                }}
              />
              {enquiryFieldErrors.phone ? <span className="field-error">{enquiryFieldErrors.phone}</span> : null}
            </label>
            <label>
              <span>Email</span>
              <input
                type="email"
                value={enquiryForm.email}
                onChange={(event) => {
                  clearEnquiryFieldError("email");
                  setEnquiryForm((current) => ({ ...current, email: event.target.value }));
                }}
              />
              {enquiryFieldErrors.email ? <span className="field-error">{enquiryFieldErrors.email}</span> : null}
            </label>
            <label>
              <span>Receiver WhatsApp Number</span>
              <input
                inputMode="numeric"
                maxLength={10}
                value={enquiryForm.receiverWhatsappNumber}
                placeholder="Optional business number"
                onChange={(event) => {
                  clearEnquiryFieldError("receiverWhatsappNumber");
                  setEnquiryForm((current) => ({
                    ...current,
                    receiverWhatsappNumber: normalizePhoneInput(event.target.value)
                  }));
                }}
              />
              {enquiryFieldErrors.receiverWhatsappNumber ? (
                <span className="field-error">{enquiryFieldErrors.receiverWhatsappNumber}</span>
              ) : null}
            </label>
            <label>
              <span>Address *</span>
              <input
                value={enquiryForm.address}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, address: event.target.value }))
                }
              />
            </label>
            <label>
              <span>Pincode *</span>
              <input
                inputMode="numeric"
                value={enquiryForm.pincode}
                onChange={(event) => {
                  clearEnquiryFieldError("pincode");
                  const value = normalizePincodeInput(event.target.value);
                  setEnquiryForm((current) => ({ ...current, pincode: value }));
                  if (value.length === 6) {
                    void autofillFromPincode(value, "source");
                  }
                }}
              />
              {enquiryFieldErrors.pincode ? <span className="field-error">{enquiryFieldErrors.pincode}</span> : null}
            </label>
            <label>
              <span>State *</span>
              <input
                value={enquiryForm.state}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, state: event.target.value }))
                }
              />
            </label>
            <label>
              <span>City *</span>
              <input
                value={enquiryForm.city}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, city: event.target.value }))
                }
              />
            </label>
            <label>
              <span>Requested asset</span>
              <input
                value={enquiryForm.requestedAsset}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, requestedAsset: event.target.value }))
                }
              />
            </label>
            <label>
              <span>Potential product</span>
              <input
                value={enquiryForm.potentialProduct}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, potentialProduct: event.target.value }))
                }
              />
            </label>
            <label className="form-span-2">
              <span>Requirement summary</span>
              <textarea
                rows={3}
                value={enquiryForm.requirementSummary}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, requirementSummary: event.target.value }))
                }
              />
            </label>
            <label className="form-span-2">
              <span className="checkbox-label">
                <input
                  type="checkbox"
                  checked={destinationSameAsMain}
                  onChange={(event) => {
                    const checked = event.target.checked;
                    setDestinationSameAsMain(checked);
                    if (checked) {
                      setEnquiryForm((current) => ({
                        ...current,
                        destinationAddress: current.address,
                        destinationPincode: current.pincode,
                        destinationState: current.state,
                        destinationCity: current.city
                      }));
                    }
                  }}
                />
                Destination same as main address
              </span>
            </label>
            <label className="form-span-2">
              <span>Destination address *</span>
              <textarea
                rows={2}
                value={enquiryForm.destinationAddress}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, destinationAddress: event.target.value }))
                }
              />
            </label>
            <label>
              <span>Destination pincode *</span>
              <input
                inputMode="numeric"
                value={enquiryForm.destinationPincode}
                onChange={(event) => {
                  clearEnquiryFieldError("destinationPincode");
                  const value = normalizePincodeInput(event.target.value);
                  setEnquiryForm((current) => ({ ...current, destinationPincode: value }));
                  if (value.length === 6) {
                    void autofillFromPincode(value, "destination");
                  }
                }}
              />
              {enquiryFieldErrors.destinationPincode ? (
                <span className="field-error">{enquiryFieldErrors.destinationPincode}</span>
              ) : null}
            </label>
            <label>
              <span>Destination state *</span>
              <input
                value={enquiryForm.destinationState}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, destinationState: event.target.value }))
                }
              />
            </label>
            <label>
              <span>Destination city *</span>
              <input
                value={enquiryForm.destinationCity}
                onChange={(event) =>
                  setEnquiryForm((current) => ({ ...current, destinationCity: event.target.value }))
                }
              />
            </label>
            </div>

            <div className="entry-actions">
              <button
                className="action-inline-button"
                type="button"
                onClick={() => void handleSubmitEnquiry()}
                disabled={isSavingEnquiry}
              >
                {isSavingEnquiry ? "Saving..." : editingEnquiryId ? "Update enquiry" : "Save enquiry"}
              </button>
            </div>
          </section>
        </section>
      );
    }

    if (entryMode === "productDocuments") {
      return (
        <section className="entry-overlay">
          <div className="entry-backdrop" onClick={closeEntryPanel} />
          <section className="panel entry-panel entry-modal" ref={entryModalRef}>
            <div className="panel-header panel-header-tight">
              <div>
                <p className="eyebrow">Product Documents</p>
                <h2>Upload and review product files</h2>
                <p className="panel-subcopy">Attach brochures, datasheets, or certificates against a product.</p>
              </div>
              <button className="entry-close" type="button" onClick={closeEntryPanel}>
                Close
              </button>
            </div>

            {popupActionState ? (
              <section className={`action-banner entry-action-banner ${popupActionState.status}`}>
                <strong>{popupActionState.label}</strong>
                <span>{popupActionState.message}</span>
                {renderBannerCloseButton(dismissActionState)}
              </section>
            ) : null}

            <div className="form-grid">
              <label className="form-span-2">
                <span>Product</span>
                <select value={selectedProductId} onChange={(event) => setSelectedProductId(event.target.value)}>
                  <option value="">Select product</option>
                  {(operations?.products ?? []).map((product) => (
                    <option key={product.id} value={product.id}>
                      {product.name || product.model || product.productKey}
                    </option>
                  ))}
                </select>
              </label>
              <label className="form-span-2">
                <span>Upload files</span>
                <input
                  type="file"
                  multiple
                  onChange={(event) => setProductDocFiles(Array.from(event.target.files || []))}
                />
              </label>
            </div>

            {selectedProduct ? (
              <div className="document-list">
                <h3>Uploaded documents for {selectedProduct.name || selectedProduct.productKey}</h3>
                {selectedProduct.documents.length ? (
                  selectedProduct.documents.map((document) => (
                    <article className="document-item" key={document.id}>
                      <div>
                        <strong>{document.fileName}</strong>
                        <span>{formatDate(document.uploadedAt)}</span>
                      </div>
                      <a className="action-inline-link" href={document.fileUrl} target="_blank" rel="noreferrer">
                        View
                      </a>
                    </article>
                  ))
                ) : (
                  <p className="entry-note">No documents uploaded for this product yet.</p>
                )}
              </div>
            ) : null}

            <div className="entry-actions">
              <p className="entry-note">Files are stored locally for now and can later be mirrored to Drive.</p>
              <button
                className="action-inline-button"
                type="button"
                onClick={() => void handleUploadProductDocuments()}
                disabled={isUploadingProductDocuments}
              >
                {isUploadingProductDocuments ? "Uploading..." : "Upload documents"}
              </button>
            </div>
          </section>
        </section>
      );
    }

    if (entryMode === "enquiryDocuments") {
      return (
        <section className="entry-overlay">
          <div className="entry-backdrop" onClick={closeEntryPanel} />
          <section className="panel entry-panel entry-modal" ref={entryModalRef}>
            <div className="panel-header panel-header-tight">
              <div>
                <p className="eyebrow">Mapped Product Documents</p>
                <h2>Review documents for this enquiry</h2>
                <p className="panel-subcopy">Select product documents and prepare them for email or WhatsApp delivery.</p>
              </div>
              <button className="entry-close" type="button" onClick={closeEntryPanel}>
                Close
              </button>
            </div>

            {popupActionState ? (
              <section className={`action-banner entry-action-banner ${popupActionState.status}`}>
                <strong>{popupActionState.label}</strong>
                <span>{popupActionState.message}</span>
                {renderBannerCloseButton(dismissActionState)}
              </section>
            ) : null}

            {selectedEnquiry?.mappedProductDocuments.length ? (
              <div className="document-list">
                {selectedEnquiry.mappedProductDocuments.map((document) => {
                  const checked = selectedDocumentIds.includes(document.id);
                  return (
                    <label className="document-item checkbox-item" key={document.id}>
                      <div className="checkbox-row">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={(event) =>
                            setSelectedDocumentIds((current) =>
                              event.target.checked
                                ? [...current, document.id]
                                : current.filter((id) => id !== document.id)
                            )
                          }
                        />
                        <div>
                          <strong>{document.fileName}</strong>
                          <span>{document.productName}</span>
                        </div>
                      </div>
                      <a className="action-inline-link" href={document.fileUrl} target="_blank" rel="noreferrer">
                        View
                      </a>
                    </label>
                  );
                })}
              </div>
            ) : (
              <p className="entry-note">No mapped product documents were found for this enquiry yet.</p>
            )}

            <div className="entry-actions">
              <p className="entry-note">
                Customer contact will be taken from the linked customer record when these are sent later.
              </p>
              <button
                className="action-inline-button"
                type="button"
                onClick={() => void handleSendProductDocuments()}
                disabled={isSendingProductDocuments}
              >
                {isSendingProductDocuments ? "Sending..." : "Send selected documents"}
              </button>
            </div>
          </section>
        </section>
      );
    }

    return (
      <section className="entry-overlay">
        <div className="entry-backdrop" onClick={closeEntryPanel} />
        <section className="panel entry-panel entry-modal" ref={entryModalRef}>
          <div className="panel-header panel-header-tight">
            <div>
              <p className="eyebrow">Portal Entry</p>
              <h2>Create quotation line items</h2>
              <p className="panel-subcopy">Select products and quantities once. The portal will create all Airtable line items.</p>
            </div>
            <button className="entry-close" type="button" onClick={closeEntryPanel}>
              Close
            </button>
          </div>

          {popupActionState ? (
            <section className={`action-banner entry-action-banner ${popupActionState.status}`}>
              <strong>{popupActionState.label}</strong>
              <span>{popupActionState.message}</span>
              {renderBannerCloseButton(dismissActionState)}
            </section>
          ) : null}

          <div className="form-grid">
            <label className="form-span-2">
              <span>Quotation</span>
              <select
                value={selectedQuotationId}
                onChange={(event) => setSelectedQuotationId(event.target.value)}
              >
                <option value="">Select quotation</option>
                {availableQuotationOptions.map((quotation) => (
                  <option key={quotation.id} value={quotation.id}>
                    {quotation.quotationNumber} - {customerLookup.get(quotation.linkedCustomerId)?.customerName || "Unlinked"}
                  </option>
                ))}
              </select>
            </label>
          </div>

          <div className="line-item-builder">
            {lineItemRows.map((row, index) => {
              const product = operations.products.find((item) => item.id === row.productId);
              const amounts = calculateLineItemAmounts(product, row.qty, row.totalAmount);

              return (
                <div className="line-item-row" key={row.id}>
                  <label className="line-item-product">
                    <span>Product {index + 1}</span>
                    <select
                      value={row.productId}
                      onChange={(event) => updateLineItemRow(row.id, "productId", event.target.value)}
                    >
                      <option value="">Select product</option>
                          {operations.products.map((productOption) => (
                            <option key={productOption.id} value={productOption.id}>
                              {productOption.name || productOption.model || productOption.productKey}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="line-item-qty">
                    <span>Qty</span>
                    <input
                      type="number"
                      min="1"
                      value={row.qty}
                      onChange={(event) => updateLineItemRow(row.id, "qty", event.target.value)}
                    />
                  </label>
                  <div className="line-item-preview">
                    <span>{product?.narration || product?.name || "Product details will auto-fill from the catalog."}</span>
                    <strong>{product ? formatCurrency(amounts.rate) : "Rate pending"}</strong>
                    {product ? (
                      <>
                        <span>{`Base ${formatCurrency(amounts.unitValue)} | Transport ${formatCurrency(amounts.transport)}`}</span>
                        <span>{`GST ${amounts.gstPercent}% = ${formatCurrency(amounts.gstAmount)}`}</span>
                      </>
                    ) : null}
                  </div>
                  <label className="line-item-total">
                    <span>Total amount</span>
                    <input
                      type="number"
                      min="0.01"
                      step="0.01"
                      value={row.totalAmount}
                      placeholder={product ? formatAmountInput(amounts.computedTotalAmount) : "Select product first"}
                      onChange={(event) => updateLineItemRow(row.id, "totalAmount", event.target.value)}
                    />
                  </label>
                  <button className="ghost-button" type="button" onClick={() => removeLineItemRow(row.id)}>
                    Remove
                  </button>
                </div>
              );
            })}
          </div>

          {selectedQuotation ? (
            <p className="entry-note">
              Adding items for <strong>{selectedQuotation.quotationNumber}</strong>.
            </p>
          ) : null}

          <div className="entry-actions">
            <button className="ghost-button" type="button" onClick={addLineItemRow}>
              Add another product
            </button>
            <button
              className="action-inline-button"
              type="button"
              onClick={() => void handleSubmitLineItems()}
              disabled={isCreatingLineItems || hasInvalidLineItems}
            >
              {isCreatingLineItems ? "Creating..." : "Create line items"}
            </button>
          </div>
        </section>
      </section>
    );
  };

  const renderDashboard = () => {
    if (!operations) {
      return null;
    }

    const growthPoints: ChartPoint[] = [
      { label: "New", value: operations.metrics.find((item) => item.label === "New Enquiries")?.value ?? 0 },
      {
        label: "Draft",
        value: operations.metrics.find((item) => item.label === "Ready For Draft")?.value ?? 0
      },
      {
        label: "Review",
        value: operations.metrics.find((item) => item.label === "Under Review")?.value ?? 0
      },
      {
        label: "Approved",
        value: operations.metrics.find((item) => item.label === "Approved")?.value ?? 0
      },
      { label: "Sent", value: operations.metrics.find((item) => item.label === "Sent")?.value ?? 0 },
      { label: "Orders", value: operations.metrics.find((item) => item.label === "Orders")?.value ?? 0 }
    ];

    const totalQuotations = operations.quotations.length;
    const sentQuotations = operations.quotations.filter((item) => item.status === "Sent").length;

    return (
      <>
        <section className="hero-card">
          <div className="hero-copy">
            <p className="eyebrow">Resham Sutra Operations</p>
            <h2>Key metrics for enquiries and orders.</h2>
          </div>
          <div className="hero-note">
            <span className="eyebrow">Quick entry</span>
            <div className="quick-actions-inline">
              <button type="button" onClick={openEnquiryEntry}>
                Create Enquiry
              </button>
              <button type="button" onClick={() => openLineItemEntry()}>
                Create Line Items
              </button>
              <a href={operations.actions.productsFormUrl || "#"} target="_blank" rel="noreferrer">
                Create Products
              </a>
            </div>
          </div>
        </section>

        {actionState && !popupActionState ? (
          <section className={`action-banner ${actionState.status}`}>
            <strong>{actionState.label}</strong>
            <span>{actionState.message}</span>
            <button type="button" className="banner-close" onClick={dismissActionState}>
              ×
            </button>
          </section>
        ) : null}

        <section className="metric-grid">
          {operations.metrics.map((metric) => (
            <article className="metric-card" key={metric.label}>
              <span>{metric.label}</span>
              <strong>{metric.value}</strong>
            </article>
          ))}
        </section>

        <section className="dashboard-grid">
          <article className="panel">
            <div className="panel-header panel-header-tight">
              <div>
                <p className="eyebrow">Activity</p>
                <h2>Attention needed</h2>
              </div>
            </div>
            <div className="feed-list">
              {dashboardFeed.length ? (
                dashboardFeed.map((item) => (
                  <article className="feed-item" key={item.id}>
                    <div>
                      <h3>{item.title}</h3>
                      <p>{item.subtitle}</p>
                    </div>
                    <div className="feed-meta">
                      <span className={`status-chip ${statusTone(item.meta)}`}>{item.meta}</span>
                      {item.href ? (
                        <a href={item.href} target="_blank" rel="noreferrer">
                          Open
                        </a>
                      ) : null}
                    </div>
                  </article>
                ))
              ) : (
                <article className="empty-card">
                  <h3>No live records</h3>
                  <p>Data will appear here once Airtable records are available.</p>
                </article>
              )}
            </div>
          </article>

          <div className="chart-stack">
            <MiniBarChart title="Workflow progression" points={growthPoints} />
            <DonutChart title="Quotations sent" value={sentQuotations} total={totalQuotations} />
          </div>
        </section>

        <section className="detail-section">
          <DetailCard
            title="Enquiry Pulse"
            rows={[
              {
                label: "Fresh enquiries",
                value: String(
                  operations.metrics.find((item) => item.label === "New Enquiries")?.value ?? 0
                )
              },
              {
                label: "Parsed",
                value: String(operations.enquiries.filter((item) => item.parserStatus === "Parsed").length)
              },
              {
                label: "Ready for draft",
                value: String(
                  operations.enquiries.filter((item) => item.parserStatus === "Ready for Draft").length
                )
              }
            ]}
          />
          <DetailCard
            title="Quotation Pulse"
            rows={[
              { label: "Draft records", value: String(operations.quotations.length) },
              {
                label: "Under review",
                value: String(
                  operations.quotations.filter((item) => item.status === "Under Review").length
                )
              },
              {
                label: "Sent",
                value: String(operations.quotations.filter((item) => item.status === "Sent").length)
              }
            ]}
          />
          <DetailCard
            title="Business Pulse"
            rows={[
              { label: "Customers", value: String(operations.customers.length) },
              { label: "Products", value: String(operations.products.length) },
              { label: "Orders", value: String(operations.orders.length) }
            ]}
          />
        </section>
      </>
    );
  };

  const renderEnquiries = () => {
    if (!operations) {
      return null;
    }

    return (
      <PaginatedTable
        eyebrow="Enquiries"
        title="All inbound requests from Airtable forms"
        rows={operations.enquiries}
        columns={
          <>
            <th>Enquiry</th>
            <th>Lead</th>
            <th>Company</th>
            <th>Status</th>
            <th>Customer</th>
            <th>Quotation</th>
            <th>Requirement</th>
            <th>Docs</th>
            <th>Folder</th>
            <th>Action</th>
          </>
        }
        emptyTitle="No enquiries yet"
        emptyBody="New Airtable form submissions will appear here."
        renderRow={(enquiry) => {
          const customer = customerLookup.get(enquiry.linkedCustomerId);
          const quotation = enquiry.quotations.length
            ? quotationLookup.get(enquiry.quotations[0])
            : undefined;

          return (
            <tr key={enquiry.id}>
              <td>
                <strong>{enquiry.enquiryId}</strong>
                <span>{[enquiry.city, enquiry.state, enquiry.pincode].filter(Boolean).join(", ") || "Address pending"}</span>
                {enquiry.loggedDateTime ? (
                  <span className="table-submeta">Logged {formatDateTime(enquiry.loggedDateTime)}</span>
                ) : null}
              </td>
              <td>
                <strong>{enquiry.leadName || "Unnamed lead"}</strong>
                <span>{enquiry.phone || enquiry.email || "Contact missing"}</span>
                {enquiry.receiverWhatsappNumber ? (
                  <span className="table-submeta">Receiver WA {enquiry.receiverWhatsappNumber}</span>
                ) : null}
              </td>
              <td>{enquiry.company || "-"}</td>
              <td>
                <span className={`status-chip ${statusTone(enquiry.parserStatus)}`}>
                  {enquiry.parserStatus || "New"}
                </span>
              </td>
              <td>{customer?.customerName || "Not linked yet"}</td>
              <td>{quotation?.quotationNumber || "Not created"}</td>
              <td>
                <strong>{enquiry.requirementSummary || "Requirement not captured"}</strong>
                <span>{enquiry.address || "Address not captured"}</span>
              </td>
              <td>
                {enquiry.mappedProductDocuments.length ? (
                  <button
                    className="action-inline-button"
                    type="button"
                    onClick={() => openEnquiryDocuments(enquiry.id)}
                  >
                    Docs ({enquiry.mappedProductDocuments.length})
                  </button>
                ) : (
                  "No docs"
                )}
              </td>
              <td>
                {enquiry.driveFolderUrl ? (
                  <a
                    className="action-inline-link"
                    href={enquiry.driveFolderUrl}
                    target="_blank"
                    rel="noreferrer"
                    onClick={() => handleLinkAction(`folder-${enquiry.id}`, "Open Folder")}
                  >
                    Open folder
                  </a>
                ) : (
                  "Not created"
                )}
              </td>
              <td>
                <div className="action-stack">
                  <button
                    className="action-inline-button"
                    type="button"
                    onClick={() => openEnquiryEdit(enquiry)}
                  >
                    Edit
                  </button>
                  {(!enquiry.linkedCustomerId || !enquiry.quotations.length) &&
                  (enquiry.parserStatus === "New" || enquiry.parserStatus === "Parsed") ? (
                    <button
                      className="action-inline-button"
                      type="button"
                      onClick={() => void handleCreateCustomer(enquiry.id)}
                      disabled={actionState?.key === enquiry.id && actionState.status === "loading"}
                    >
                      {actionState?.key === enquiry.id && actionState.status === "loading"
                        ? "Creating..."
                        : "Create Customer"}
                    </button>
                  ) : enquiry.parserStatus === "Parsed" &&
                    enquiry.linkedCustomerId &&
                    enquiry.quotations.length ? (
                    <button
                      className="action-inline-button"
                      type="button"
                      onClick={() => {
                        openLineItemEntry(quotation?.id || "");
                        setActionState({
                          key: `line-items-${enquiry.id}`,
                          label: quotation?.quotationNumber
                            ? `Add Line Items for ${quotation.quotationNumber}`
                            : "Add Line Items",
                          status: "success",
                          message: "Line item entry is ready in the portal."
                        });
                      }}
                    >
                      {quotation?.quotationNumber ? `Add Items (${quotation.quotationNumber})` : "Add Line Items"}
                    </button>
                  ) : enquiry.parserStatus === "Ready for Review" && enquiry.driveFolderUrl ? (
                    <a
                      className="action-inline-link"
                      href={enquiry.driveFolderUrl}
                      target="_blank"
                      rel="noreferrer"
                      onClick={() => handleLinkAction(`review-${enquiry.id}`, "Open Folder")}
                    >
                      Open Folder
                    </a>
                  ) : (
                    <span>Auto</span>
                  )}
                </div>
              </td>
            </tr>
          );
        }}
      />
    );
  };

  const renderCustomers = () => {
    if (!operations) {
      return null;
    }

    return (
      <PaginatedTable
        eyebrow="Customers"
        title="Master account records generated from enquiry intake"
        rows={operations.customers}
        columns={
          <>
            <th>Client ID</th>
            <th>Customer</th>
            <th>Company</th>
            <th>Contact</th>
            <th>Type</th>
            <th>Drive</th>
          </>
        }
        emptyTitle="No customers yet"
        emptyBody="Customer records created from enquiries will appear here."
        renderRow={(customer) => (
          <tr key={customer.id}>
            <td>{customer.clientId || "Pending"}</td>
            <td>
              <strong>{customer.customerName || "Unnamed customer"}</strong>
              <span>{customer.email || "Email not set"}</span>
            </td>
            <td>{customer.company || "-"}</td>
            <td>
              <strong>{customer.phone || "No phone"}</strong>
              <span>{[customer.city, customer.state, customer.pincode].filter(Boolean).join(", ") || customer.address || "Address pending"}</span>
            </td>
            <td>{customer.customerType || "Unclassified"}</td>
            <td>
              {customer.driveFolderUrl ? (
                <a href={customer.driveFolderUrl} target="_blank" rel="noreferrer">
                  Open folder
                </a>
              ) : (
                "Pending"
              )}
            </td>
          </tr>
        )}
      />
    );
  };

  const renderQuotationActionButtons = (quotation: QuotationRecord) => {
    const emailActionKey = `quotation-send-email-${quotation.id}`;
    const whatsappActionKey = `quotation-send-whatsapp-${quotation.id}`;
    const regenerateActionKey = `quotation-regenerate-${quotation.id}`;
    const isEmailLoading =
      actionState?.key === emailActionKey && actionState.status === "loading";
    const isWhatsAppLoading =
      actionState?.key === whatsappActionKey && actionState.status === "loading";
    const isRegenerateLoading =
      actionState?.key === regenerateActionKey && actionState.status === "loading";

    if (quotation.status === "Parsed") {
      return (
        <button
          className="action-inline-button"
          type="button"
          onClick={() => {
            openLineItemEntry(quotation.id);
            setActionState({
              key: `quotation-line-items-${quotation.id}`,
              label: `Add Line Items for ${quotation.quotationNumber}`,
              status: "success",
              message: "Line item entry is ready in the portal."
            });
          }}
        >
          Add Line Items
        </button>
      );
    }

    if (quotation.status === "Approved") {
      return (
        <div className="action-stack">
          <div className="action-icon-row">
            <button
              className="icon-action-button neutral"
              type="button"
              title="Regenerate final PDF"
              onClick={() => void handleGenerateFinalPdf(quotation.id)}
              disabled={
                actionState?.key === `pdf-generate-${quotation.id}` &&
                actionState.status === "loading"
              }
            >
              {actionState?.key === `pdf-generate-${quotation.id}` &&
              actionState.status === "loading"
                ? "..."
                : "PDF"}
            </button>
            <button
              className="icon-action-button"
              type="button"
              title="Send quotation on email"
              onClick={() => void handleSendQuotation(quotation.id, "email")}
              disabled={isEmailLoading}
            >
              {isEmailLoading ? "..." : "@"}
            </button>
            <button
              className="icon-action-button whatsapp"
              type="button"
              title="Send quotation on WhatsApp"
              onClick={() => void handleSendQuotation(quotation.id, "whatsapp")}
              disabled={isWhatsAppLoading}
            >
              {isWhatsAppLoading ? "..." : "WA"}
            </button>
          </div>
          {quotation.driveFolderUrl ? (
            <a
              className="action-inline-link"
              href={quotation.driveFolderUrl}
              target="_blank"
              rel="noreferrer"
              onClick={() => handleLinkAction(`quotation-folder-${quotation.id}`, "Open Folder")}
            >
              Open Folder
            </a>
          ) : null}
        </div>
      );
    }

    if (quotation.status === "Draft Sent" || quotation.status === "Sent") {
      return (
        <div className="action-stack">
          <div className="action-icon-row">
            <button
              className="icon-action-button"
              type="button"
              title="Send quotation again on email"
              onClick={() => void handleSendQuotation(quotation.id, "email")}
              disabled={isEmailLoading}
            >
              {isEmailLoading ? "..." : "@"}
            </button>
            <button
              className="icon-action-button whatsapp"
              type="button"
              title="Send quotation again on WhatsApp"
              onClick={() => void handleSendQuotation(quotation.id, "whatsapp")}
              disabled={isWhatsAppLoading}
            >
              {isWhatsAppLoading ? "..." : "WA"}
            </button>
            <button
              className="icon-action-button neutral"
              type="button"
              title="Regenerate quotation draft"
              onClick={() => void handleRegenerateQuotationDraft(quotation.id)}
              disabled={isRegenerateLoading}
            >
              {isRegenerateLoading ? "..." : "R"}
            </button>
          </div>
          {quotation.driveFolderUrl ? (
            <a
              className="action-inline-link"
              href={quotation.driveFolderUrl}
              target="_blank"
              rel="noreferrer"
              onClick={() => handleLinkAction(`quotation-folder-${quotation.id}`, "Open Folder")}
            >
              Open Folder
            </a>
          ) : null}
        </div>
      );
    }

    if (quotation.draftFileUrl || quotation.status === "Ready for Review") {
      return (
        <div className="action-stack">
          <button
            className="action-inline-button"
            type="button"
            onClick={() => void handleGenerateFinalPdf(quotation.id)}
            disabled={
              actionState?.key === `pdf-generate-${quotation.id}` &&
              actionState.status === "loading"
            }
          >
            {actionState?.key === `pdf-generate-${quotation.id}` &&
            actionState.status === "loading"
              ? "Generating..."
              : "Generate Final PDF"}
          </button>
          {quotation.driveFolderUrl ? (
            <a
              className="action-inline-link"
              href={quotation.driveFolderUrl}
              target="_blank"
              rel="noreferrer"
              onClick={() => handleLinkAction(`quotation-folder-${quotation.id}`, "Open Folder")}
            >
              Open Folder
            </a>
          ) : null}
        </div>
      );
    }

    return <span>{formatDate(quotation.sentDate)}</span>;
  };

  const renderQuotations = (title: string, subtitle: string, statuses: string[]) => {
    if (!operations) {
      return null;
    }

    const filteredQuotations = operations.quotations.filter((quotation) =>
      statuses.length ? statuses.includes(quotation.status) : true
    );
    const quotationActionState =
      actionState &&
      (actionState.key.startsWith("quotation-send-") ||
        actionState.key.startsWith("quotation-regenerate-") ||
        actionState.key.startsWith("pdf-generate-"))
        ? actionState
        : null;

    return (
      <>
        {quotationActionState ? (
          <section className={`action-banner ${quotationActionState.status}`}>
            <strong>{quotationActionState.label}</strong>
            <span>{quotationActionState.message}</span>
            <button type="button" className="banner-close" onClick={dismissActionState}>
              ×
            </button>
          </section>
        ) : null}
        <PaginatedTable
          eyebrow="Quotations"
          title={title}
          subtitle={subtitle}
          rows={filteredQuotations}
          columns={
            <>
              <th>Quotation</th>
              <th>Customer</th>
              <th>Status</th>
              <th>Channel</th>
              <th>Items</th>
              {statuses.length === 1 && statuses[0] === "Approved" ? null : <th>Draft</th>}
              <th>Final PDF</th>
              <th>Actions</th>
            </>
          }
          emptyTitle="No quotations in this section"
          emptyBody="As the workflow progresses, matching quotation records will appear here."
          renderRow={(quotation) => (
            <tr key={quotation.id}>
              <td>
                <strong>{quotation.quotationNumber}</strong>
                <span>{enquiryLookup.get(quotation.linkedEnquiryId)?.enquiryId || "No enquiry link"}</span>
                {quotation.loggedDateTime ? (
                  <span className="table-submeta">Logged {formatDateTime(quotation.loggedDateTime)}</span>
                ) : null}
              </td>
              <td>{customerLookup.get(quotation.linkedCustomerId)?.customerName || "Not linked"}</td>
              <td>
                <span className={`status-chip ${statusTone(quotation.status)}`}>
                  {quotation.status || "Draft"}
                </span>
              </td>
              <td>
                <div className="table-link-stack">
                  <span>{quotation.preferredSendChannel || "Email"}</span>
                  {quotation.emailSentDateTime ? (
                    <span className="table-submeta">
                      Email {formatDateTime(quotation.emailSentDateTime)}
                    </span>
                  ) : null}
                  {quotation.whatsappSentDateTime ? (
                    <span className="table-submeta">
                      WhatsApp {formatDateTime(quotation.whatsappSentDateTime)}
                    </span>
                  ) : null}
                </div>
              </td>
              <td>{quotation.lineItemCount || 0}</td>
              {statuses.length === 1 && statuses[0] === "Approved" ? null : (
                <td>
                  {quotation.status === "Approved" ? (
                    quotation.driveFolderUrl ? (
                      <a
                        className="action-inline-link"
                        href={quotation.driveFolderUrl}
                        target="_blank"
                        rel="noreferrer"
                        onClick={() => handleLinkAction(`quotation-folder-${quotation.id}`, "Open Folder")}
                      >
                        Open Folder
                      </a>
                    ) : (
                      "-"
                    )
                  ) : quotation.draftFileUrl ? (
                    <a
                      className="action-inline-link"
                      href={quotation.draftFileUrl}
                      target="_blank"
                      rel="noreferrer"
                      onClick={() => handleLinkAction(`draft-${quotation.id}`, "Open Draft")}
                    >
                      Open draft
                    </a>
                  ) : (
                    "Not generated"
                  )}
                </td>
              )}
              <td>
                {quotation.finalPdfUrl ? (
                  <div className="table-link-stack">
                    <a
                      className="action-inline-link"
                      href={quotation.finalPdfUrl}
                      target="_blank"
                      rel="noreferrer"
                      onClick={() => handleLinkAction(`pdf-${quotation.id}`, "Open PDF")}
                    >
                      Open PDF
                    </a>
                    {quotation.finalPdfGeneratedAt ? (
                      <span className="table-submeta">
                        Last generated {formatDateTime(quotation.finalPdfGeneratedAt)}
                      </span>
                    ) : null}
                    {quotation.draftCreatedTime ? (
                      <span className="table-submeta">
                        Draft created {formatDateTime(quotation.draftCreatedTime)}
                      </span>
                    ) : null}
                  </div>
                ) : (
                  quotation.draftCreatedTime ? (
                    <span className="table-submeta">
                      Draft created {formatDateTime(quotation.draftCreatedTime)}
                    </span>
                  ) : (
                    "Pending"
                  )
                )}
              </td>
              <td>{renderQuotationActionButtons(quotation)}</td>
            </tr>
          )}
        />
      </>
    );
  };

  const renderOrders = () => {
    if (!operations) {
      return null;
    }

    return (
      <PaginatedTable
        eyebrow="Orders"
        title="Converted business once quotations are accepted"
        rows={operations.orders}
        columns={
          <>
            <th>Order</th>
            <th>Customer</th>
            <th>Quotation</th>
            <th>Order Date</th>
            <th>Value</th>
            <th>Payment</th>
            <th>Delivery</th>
          </>
        }
        emptyTitle="No orders yet"
        emptyBody="Accepted quotations will surface here as orders."
        renderRow={(order) => (
          <tr key={order.id}>
            <td>{order.orderNumber}</td>
            <td>{customerLookup.get(order.linkedCustomerId)?.customerName || "Not linked"}</td>
            <td>{quotationLookup.get(order.linkedQuotationId)?.quotationNumber || "Not linked"}</td>
            <td>{formatDate(order.orderDate)}</td>
            <td>{order.orderValue ? formatCurrency(order.orderValue) : "Pending"}</td>
            <td>{order.paymentStatus || "Pending"}</td>
            <td>{order.deliveryStatus || "Pending"}</td>
          </tr>
        )}
      />
    );
  };

  const renderProducts = () => {
    if (!operations) {
      return null;
    }

    return (
      <PaginatedTable
        eyebrow="Products"
        title="Catalog and pricing references feeding quotation line item selection"
        rows={operations.products}
        columns={
          <>
            <th>Product Key</th>
            <th>Model</th>
            <th>Name</th>
            <th>Narration</th>
            <th>Bulk Sale</th>
            <th>MRP</th>
            <th>Source</th>
            <th>Documents</th>
          </>
        }
        emptyTitle="No products loaded"
        emptyBody="Catalog rows will appear here after product entries are added."
        renderRow={(product) => (
          <tr key={product.id}>
            <td>{product.productKey}</td>
            <td>{product.model || "-"}</td>
            <td>{product.name || "Unnamed product"}</td>
            <td>{product.narration || "No narration"}</td>
            <td>{product.bulkSalePrice ? formatCurrency(product.bulkSalePrice) : "Pending"}</td>
            <td>{product.mrp ? formatCurrency(product.mrp) : "Pending"}</td>
            <td>{product.sourceSheet || "Manual"}</td>
            <td>
              <div className="action-stack">
                <button
                  className="action-inline-button"
                  type="button"
                  onClick={() => openProductDocuments(product.id)}
                >
                  Upload / View ({product.documents.length})
                </button>
                {product.documents[0] ? (
                  <a
                    className="action-inline-link"
                    href={product.documents[0].fileUrl}
                    target="_blank"
                    rel="noreferrer"
                  >
                    Latest file
                  </a>
                ) : null}
              </div>
            </td>
          </tr>
        )}
      />
    );
  };

  if (authLoading) {
    return (
      <main className="auth-shell">
        <section className="auth-card auth-card-loading">
          <img className="auth-logo" src={logoUrl} alt="Resham Sutra" />
          <p className="eyebrow">Secure Access</p>
          <h1>Loading your workspace</h1>
          <p>Checking for an active ReshamSutra session.</p>
        </section>
      </main>
    );
  }

  if (!currentUser) {
    return (
      <main className="auth-shell">
        <section className="auth-card">
          <img className="auth-logo" src={logoUrl} alt="Resham Sutra" />
          <p className="eyebrow">Secure Access</p>
          <h1>Sign in to ReshamSutra Operations</h1>
          <p className="auth-copy">
            Only users listed in the Airtable <strong>ReshamSutra Users</strong> table can access
            the live dashboard.
          </p>
          {authActionState ? (
            <section className={`action-banner auth-banner ${authActionState.status}`}>
              <strong>{authActionState.label}</strong>
              <span>{authActionState.message}</span>
              <button type="button" className="banner-close" onClick={dismissAuthActionState}>
                ×
              </button>
            </section>
          ) : null}
          <div className="auth-form">
            <label>
              <span>Email</span>
              <input
                type="email"
                autoComplete="username"
                value={loginForm.email}
                onChange={(event) =>
                  setLoginForm((current) => ({
                    ...current,
                    email: event.target.value
                  }))
                }
              />
            </label>
            <label>
              <span>Password</span>
              <input
                type="password"
                autoComplete="current-password"
                value={loginForm.password}
                onChange={(event) =>
                  setLoginForm((current) => ({
                    ...current,
                    password: event.target.value
                  }))
                }
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    void handleLogin();
                  }
                }}
              />
            </label>
          </div>
          {authError ? <p className="auth-error">{authError}</p> : null}
          <div className="auth-actions">
            <span>Your login stays active on this browser for 30 days.</span>
            <button
              type="button"
              className="action-inline-button"
              onClick={() => void handleLogin()}
              disabled={authActionState?.key === "auth-login" && authActionState.status === "loading"}
            >
              {authActionState?.key === "auth-login" && authActionState.status === "loading"
                ? "Signing in..."
                : "Sign In"}
            </button>
          </div>
        </section>
      </main>
    );
  }

  let content: ReactNode = null;

  if (loading) {
    content = (
      <section className="panel loading-panel">
        <p className="eyebrow">Loading</p>
        <h2>Pulling live Airtable operations data into the dashboard.</h2>
      </section>
    );
  } else if (error) {
    content = (
      <section className="panel loading-panel">
        <p className="eyebrow">Connection issue</p>
        <h2>We could not load the operations snapshot.</h2>
        <p>{error}</p>
      </section>
    );
  } else {
    switch (activeView) {
      case "dashboard":
        content = renderDashboard();
        break;
      case "orders":
        content = renderOrders();
        break;
      case "sentQuotations":
        content = renderQuotations(
          "Sent quotations and follow-through visibility",
          "Track what has already gone out to clients and when it was sent.",
          ["Sent", "Draft Sent"]
        );
        break;
      case "approvedQuotations":
        content = renderQuotations(
          "Approved quotations ready for dispatch",
          "Quotes that have cleared internal review and are waiting for send or final conversion.",
          ["Approved"]
        );
        break;
      case "quotationDrafts":
        content = renderQuotations(
          "Quotation drafts prepared from live enquiries",
          "Internal working drafts that still need line item and commercial review.",
          ["Draft", "Under Review", "Ready for Draft", "Ready for Review", "Parsed"]
        );
        break;
      case "customers":
        content = renderCustomers();
        break;
      case "enquiries":
        content = renderEnquiries();
        break;
      case "products":
        content = renderProducts();
        break;
      default:
        content = renderDashboard();
    }
  }

  return (
    <main className={sidebarCollapsed ? "app-shell sidebar-collapsed" : "app-shell"}>
      <aside className="sidebar">
        <div className="brand-card">
          <div className="brand-row">
            <img className="brand-logo" src={logoUrl} alt="Resham Sutra" />
            {!sidebarCollapsed ? (
              <div className="brand-text">
                <h1>ReshamSutra</h1>
                <span>Operations</span>
              </div>
            ) : null}
          </div>
        </div>

        <nav className="nav-stack">
          {viewOptions.map((view) => (
            <button
              key={`${view.key}-${view.label}`}
              className={activeView === view.key ? "nav-item active" : "nav-item"}
              onClick={() => setActiveView(view.key)}
              type="button"
              title={view.label}
            >
              <span>{sidebarCollapsed ? view.label.charAt(0) : view.label}</span>
            </button>
          ))}
        </nav>

        <div className="sidebar-footer">
          <button
            className="collapse-button"
            type="button"
            onClick={() => setSidebarCollapsed((current) => !current)}
          >
            {sidebarCollapsed ? ">" : "<"}
          </button>
        </div>
      </aside>

      <section className="workspace">
        <header className="workspace-header">
          <div className="workspace-copy">
            <div className="breadcrumbs">
              <span>Resham Sutra Operations</span>
              <span className="crumb-sep">›</span>
              <strong>{viewOptions.find((view) => view.key === activeView)?.label ?? "Dashboard"}</strong>
            </div>
          </div>
          <div className="workspace-tools" ref={profileMenuRef}>
            <button className="logout-button" type="button" onClick={() => void handleLogout()}>
              Logout
            </button>
            <button
              className="profile-button"
              type="button"
              title={`${currentUser.name} • ${currentUser.email}`}
              onClick={() => setProfileOpen((current) => !current)}
            >
              {currentUser.name
                .split(/\s+/)
                .filter(Boolean)
                .slice(0, 2)
                .map((part) => part.charAt(0).toUpperCase())
                .join("") || "RS"}
            </button>
            {profileOpen ? (
              <div className="profile-card">
                <strong>{currentUser.name}</strong>
                <span>{currentUser.email}</span>
              </div>
            ) : null}
          </div>
        </header>

        {renderEntryPanel()}
        {content}
      </section>
    </main>
  );
}
