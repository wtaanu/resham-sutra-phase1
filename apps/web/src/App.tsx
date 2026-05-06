import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import logoFallbackUrl from "./assets/resham-sutra-logo.png";
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
  defaultTemplateFolderUrl: string;
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
  potentialProduct: string;
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
  referenceNumber: string;
  loggedDateTime: string;
  linkedCustomerId: string;
  linkedEnquiryId: string;
  status: string;
  draftFileUrl: string;
  draftCreatedTime: string;
  finalPdfUrl: string;
  driveFolderUrl: string;
  sentDate: string;
  whatsappSentDateTime: string;
  emailSentDateTime: string;
  finalPdfGeneratedAt?: string;
  lineItemCount?: number;
  quotationGrandTotal?: number;
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
  linkedEnquiryId?: string;
  orderDate: string;
  orderStatus?: string;
  totalAmount?: number;
  orderNotes?: string;
  quotationGrandTotal?: number;
  quotationStatus?: string;
  orderLineItemCount?: number;
  orderValuePerItem?: string;
  orderFulfillmentProgress?: string;
  orderSummary?: string;
  orderRiskAttentionFlag?: string;
  orderValue: number;
  paymentStatus: string;
  deliveryStatus: string;
  address: string;
  state: string;
  city: string;
  pincode: string;
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
  totals?: {
    enquiries: number;
    customers: number;
    quotations: number;
    orders: number;
    products: number;
  };
  enquiries: EnquiryRecord[];
  customers: CustomerRecord[];
  quotations: QuotationRecord[];
  orders: OrderRecord[];
  products: ProductRecord[];
};

type CustomersPageResponse = {
  status: string;
  customers: CustomerRecord[];
  nextOffset: string;
  pageSize: number;
  totalCount: number;
};

type EnquiriesPageResponse = {
  status: string;
  enquiries: EnquiryRecord[];
  nextOffset: string;
  pageSize: number;
  totalCount: number;
};

type QuotationsPageResponse = {
  status: string;
  quotations: QuotationRecord[];
  nextOffset: string;
  pageSize: number;
  totalCount: number;
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

type EntryMode =
  | "enquiry"
  | "customer"
  | "quotation"
  | "order"
  | "lineItems"
  | "productDocuments"
  | "enquiryDocuments"
  | null;

type EnquiryFormState = {
  linkedCustomerId: string;
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
  requirementSummary: string;
  potentialProduct: string;
};

type EnquiryFieldErrors = Partial<Record<keyof EnquiryFormState, string>>;

type CustomerFormState = {
  customerName: string;
  company: string;
  phone: string;
  email: string;
  address: string;
  state: string;
  city: string;
  pincode: string;
  customerType: string;
};

type QuotationFormState = {
  enquiryId: string;
};

type OrderFormState = {
  quotationId: string;
  customerId: string;
  enquiryId: string;
  orderDate: string;
  orderStatus: "Confirmed" | "Processing" | "Shipped" | "Delivered" | "Cancelled";
  totalAmount: string;
  orderNotes: string;
  paymentStatus: "Paid" | "Pending" | "Half Payment";
  address: string;
  state: string;
  city: string;
  pincode: string;
};

type OrderLineItemRow = {
  id: string;
  productId: string;
  description: string;
  qty: string;
  ratePerUnit: string;
  packingFreight: string;
  unitValue: string;
  gst18: string;
  totalAmount: string;
};

type LineItemDraftRow = {
  id: string;
  productId: string;
  qty: string;
  rate: string;
  transport: string;
  gstPercent: string;
  existing: boolean;
};

type PortalLineItemResponse = {
  id: string;
  productId: string;
  qty: number;
  rate: number;
  transport: number;
  gstPercent: number;
  totalAmount: number;
};

type PortalOrderLineItemResponse = {
  id: string;
  productId: string;
  description: string;
  qty: number;
  ratePerUnit: number;
  packingFreight: number;
  unitValue: number;
  gst18: number;
  totalAmount: number;
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
  headerAction?: ReactNode;
  rows: T[];
  columns: ReactNode;
  renderRow: (row: T) => ReactNode;
  emptyTitle: string;
  emptyBody: string;
  pageSize?: number;
  pagination?: {
    canNext: boolean;
    canPrevious: boolean;
    isLoading?: boolean;
    label: string;
    metaLabel?: string;
    onNext: () => void;
    onPrevious: () => void;
  };
};

const apiUrl = import.meta.env.VITE_API_URL ?? "http://localhost:4000";

function ReshamSutraLogo({ className }: { className: string }) {
  const [src, setSrc] = useState(logoUrl);

  if (!src) {
    return (
      <span className={`${className} logo-fallback`} aria-label="Resham Sutra">
        RS
      </span>
    );
  }

  return (
    <img
      className={className}
      src={src}
      alt="Resham Sutra"
      decoding="async"
      onError={() => setSrc((current) => (current === logoUrl ? logoFallbackUrl : ""))}
    />
  );
}

const viewOptions: NavView[] = [
  { key: "dashboard", label: "Dashboard" },
  { key: "orders", label: "Orders" },
  { key: "sentQuotations", label: "Sent Quotations" },
  { key: "approvedQuotations", label: "Approved Quotations" },
  { key: "quotationDrafts", label: "Quotation Drafts" },
  { key: "customers", label: "Customers" },
  { key: "enquiries", label: "Enquiries" },
  { key: "products", label: "Products" }
];

const enquiryWorkflowStatuses = [
  "New",
  "New Enquiries",
  "Parsed",
  "Draft Quote",
  "Approved Quote",
  "Sent Quote",
  "Ordered"
];

const stateOptions = [
  "Andhra Pradesh",
  "Arunachal Pradesh",
  "Assam",
  "Bihar",
  "Chhattisgarh",
  "Delhi",
  "Goa",
  "Gujarat",
  "Haryana",
  "Himachal Pradesh",
  "Jharkhand",
  "Karnataka",
  "Kerala",
  "Madhya Pradesh",
  "Maharashtra",
  "Manipur",
  "Meghalaya",
  "Mizoram",
  "Nagaland",
  "Odisha",
  "Punjab",
  "Rajasthan",
  "Sikkim",
  "Tamil Nadu",
  "Telangana",
  "Tripura",
  "Uttar Pradesh",
  "Uttarakhand",
  "West Bengal"
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

function firstEnquiryFieldError(errors: EnquiryFieldErrors) {
  return Object.values(errors).find(Boolean) || "Please fix the highlighted fields.";
}

function calculateLineItemAmounts(
  product: ProductRecord | undefined,
  qtyInput: string | number,
  rateInput?: string | number,
  transportInput?: string | number,
  gstPercentInput?: string | number
) {
  const qty = Math.max(1, Number(qtyInput || 0));
  const resolvedRate =
    rateInput !== undefined && String(rateInput).trim() !== ""
      ? Number(rateInput)
      : Number(product?.bulkSalePrice || product?.mrp || 0);
  const resolvedTransport =
    transportInput !== undefined && String(transportInput).trim() !== ""
      ? Number(transportInput)
      : Number(product?.transportCharge || 0);
  const resolvedGstPercent =
    gstPercentInput !== undefined && String(gstPercentInput).trim() !== ""
      ? Number(gstPercentInput)
      : Number(product?.gstPercent || 0);
  const rate = Number.isFinite(resolvedRate) ? resolvedRate : 0;
  const transport = Number.isFinite(resolvedTransport) ? resolvedTransport : 0;
  const gstPercent = Number.isFinite(resolvedGstPercent) ? resolvedGstPercent : 0;
  const unitValue = Number((rate * qty).toFixed(2));
  const freightAmount = Number((transport * qty).toFixed(2));
  const computedGstAmount = Number((gstPercent * qty).toFixed(2));
  const computedTotalAmount = Number((unitValue + freightAmount + computedGstAmount).toFixed(2));

  return {
    qty,
    rate,
    transport,
    gstPercent,
    unitValue,
    freightAmount,
    gstAmount: computedGstAmount,
    totalAmount: computedTotalAmount,
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
    address: "",
    state: "",
    city: "",
    pincode: "",
    destinationAddress: "",
    destinationState: "",
    destinationCity: "",
    destinationPincode: "",
    requirementSummary: "",
    potentialProduct: ""
  };
}

function createLineItemRow(): LineItemDraftRow {
  return {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    productId: "",
    qty: "1",
    rate: "",
    transport: "",
    gstPercent: "",
    existing: false
  };
}

function createOrderLineItemRow(): OrderLineItemRow {
  return {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    productId: "",
    description: "",
    qty: "1",
    ratePerUnit: "",
    packingFreight: "",
    unitValue: "",
    gst18: "",
    totalAmount: ""
  };
}

function createBlankCustomerForm(): CustomerFormState {
  return {
    customerName: "",
    company: "",
    phone: "",
    email: "",
    address: "",
    state: "",
    city: "",
    pincode: "",
    customerType: "Domestic"
  };
}

function createBlankQuotationForm(): QuotationFormState {
  return {
    enquiryId: ""
  };
}

function createBlankOrderForm(): OrderFormState {
  return {
    quotationId: "",
    customerId: "",
    enquiryId: "",
    orderDate: "",
    orderStatus: "Confirmed",
    totalAmount: "",
    orderNotes: "",
    paymentStatus: "Pending",
    address: "",
    state: "",
    city: "",
    pincode: ""
  };
}

function mapPortalLineItemToDraftRow(item: PortalLineItemResponse): LineItemDraftRow {
  return {
    id: item.id || `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    productId: item.productId,
    qty: String(item.qty || 1),
    rate: formatAmountInput(item.rate || 0),
    transport: formatAmountInput(item.transport || 0),
    gstPercent: formatAmountInput(item.gstPercent || 0),
    existing: true
  };
}

function mapPortalOrderLineItemToDraftRow(item: PortalOrderLineItemResponse): OrderLineItemRow {
  return {
    id: item.id || `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    productId: item.productId || "",
    description: item.description || "",
    qty: String(item.qty || 1),
    ratePerUnit: formatAmountInput(item.ratePerUnit || 0),
    packingFreight: formatAmountInput(item.packingFreight || 0),
    unitValue: formatAmountInput(item.unitValue || 0),
    gst18: formatAmountInput(item.gst18 || 0),
    totalAmount: formatAmountInput(item.totalAmount || 0)
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
  headerAction,
  pageSize = 8,
  pagination
}: PaginatedTableProps<T>) {
  const [page, setPage] = useState(1);
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));

  useEffect(() => {
    setPage(1);
  }, [rows.length, title]);

  const start = (page - 1) * pageSize;
  const pageRows = pagination ? rows : rows.slice(start, start + pageSize);

  return (
    <section className="panel">
      <div className="panel-header panel-header-tight">
        <div>
          <p className="eyebrow">{eyebrow}</p>
          <h2>{title}</h2>
          {subtitle ? <p className="panel-subcopy">{subtitle}</p> : null}
        </div>
        <div className="table-header-actions">
          {headerAction}
          <span className="table-meta">{pagination?.metaLabel || `${rows.length} records`}</span>
        </div>
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
              {pagination?.label || `Page ${page} of ${totalPages}`}
            </span>
            <div className="pagination-actions">
              <button
                type="button"
                disabled={pagination ? !pagination.canPrevious || pagination.isLoading : page === 1}
                onClick={() => (pagination ? pagination.onPrevious() : setPage(page - 1))}
              >
                Previous
              </button>
              <button
                type="button"
                disabled={pagination ? !pagination.canNext || pagination.isLoading : page === totalPages}
                onClick={() => (pagination ? pagination.onNext() : setPage(page + 1))}
              >
                {pagination?.isLoading ? "Loading..." : "Next"}
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
  const createPagedState = <T,>() => ({
    currentOffset: "",
    error: "",
    initialized: false,
    loading: false,
    nextOffset: "",
    page: 1,
    previousOffsets: [] as string[],
    rows: [] as T[],
    totalCount: 0
  });
  const [enquiryPage, setEnquiryPage] = useState(createPagedState<EnquiryRecord>);
  const [customerPage, setCustomerPage] = useState<{
    currentOffset: string;
    error: string;
    initialized: boolean;
    loading: boolean;
    nextOffset: string;
    page: number;
    previousOffsets: string[];
    rows: CustomerRecord[];
    totalCount: number;
  }>(createPagedState<CustomerRecord>);
  const [quotationPage, setQuotationPage] = useState(createPagedState<QuotationRecord>);
  const [quotationPageKey, setQuotationPageKey] = useState("");
  const [activeView, setActiveView] = useState<ViewKey>("dashboard");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [actionState, setActionState] = useState<ActionState | null>(null);
  const [entryMode, setEntryMode] = useState<EntryMode>(null);
  const [editingEnquiryId, setEditingEnquiryId] = useState("");
  const [editingCustomerId, setEditingCustomerId] = useState("");
  const [editingOrderId, setEditingOrderId] = useState("");
  const [enquiryForm, setEnquiryForm] = useState<EnquiryFormState>(createBlankEnquiryForm);
  const [enquiryFieldErrors, setEnquiryFieldErrors] = useState<EnquiryFieldErrors>({});
  const [customerForm, setCustomerForm] = useState<CustomerFormState>(createBlankCustomerForm);
  const [quotationForm, setQuotationForm] = useState<QuotationFormState>(createBlankQuotationForm);
  const [orderForm, setOrderForm] = useState<OrderFormState>(createBlankOrderForm);
  const [orderLineItems, setOrderLineItems] = useState<OrderLineItemRow[]>([createOrderLineItemRow()]);
  const [destinationSameAsMain, setDestinationSameAsMain] = useState(false);
  const [customerSearchTerm, setCustomerSearchTerm] = useState("");
  const [customerDropdownOpen, setCustomerDropdownOpen] = useState(false);
  const [productSearchTerm, setProductSearchTerm] = useState("");
  const [productDropdownOpen, setProductDropdownOpen] = useState(false);
  const [enquiryStatusFilter, setEnquiryStatusFilter] = useState("All");
  const [selectedQuotationId, setSelectedQuotationId] = useState("");
  const [lineItemRows, setLineItemRows] = useState<LineItemDraftRow[]>([createLineItemRow()]);
  const [selectedProductId, setSelectedProductId] = useState("");
  const [productDocFiles, setProductDocFiles] = useState<File[]>([]);
  const [selectedEnquiryId, setSelectedEnquiryId] = useState("");
  const [selectedDocumentIds, setSelectedDocumentIds] = useState<string[]>([]);
  const entryModalRef = useRef<HTMLElement | null>(null);
  const profileMenuRef = useRef<HTMLDivElement | null>(null);
  const enquirySubmitInFlightRef = useRef(false);
  const customerSubmitInFlightRef = useRef(false);
  const quotationSubmitInFlightRef = useRef(false);
  const orderSubmitInFlightRef = useRef(false);
  const lineItemsSubmitInFlightRef = useRef(false);

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

  function quotationStatusesForView(view: ViewKey) {
    if (view === "sentQuotations") {
      return ["Sent Quote"];
    }
    if (view === "approvedQuotations") {
      return ["Approved Quote"];
    }
    if (view === "quotationDrafts") {
      return ["Draft Quote", "Parsed", "New Enquiries"];
    }
    return [];
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
      setEnquiryPage(createPagedState<EnquiryRecord>());
      setCustomerPage((current) => ({ ...current, initialized: false, rows: [] }));
      setQuotationPage(createPagedState<QuotationRecord>());
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
      setEnquiryPage(createPagedState<EnquiryRecord>());
      setCustomerPage((current) => ({ ...current, initialized: false, rows: [] }));
      setQuotationPage(createPagedState<QuotationRecord>());
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
    }, 60000);

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
    if (activeView !== "customers" || !currentUser || customerPage.initialized || customerPage.loading) {
      return;
    }

    void loadCustomerPage("", "reset");
  }, [activeView, currentUser, customerPage.initialized, customerPage.loading]);

  useEffect(() => {
    if (activeView !== "enquiries" || !currentUser || enquiryPage.initialized || enquiryPage.loading) {
      return;
    }

    void loadEnquiryPage("", "reset");
  }, [activeView, currentUser, enquiryPage.initialized, enquiryPage.loading, enquiryStatusFilter]);

  useEffect(() => {
    const statuses = quotationStatusesForView(activeView);
    const key = statuses.join("|") || "all";
    if (!statuses.length || !currentUser) {
      return;
    }

    if (quotationPageKey && quotationPageKey !== key) {
      setQuotationPage(createPagedState<QuotationRecord>());
      setQuotationPageKey("");
      return;
    }

    if (!quotationPage.initialized && !quotationPage.loading) {
      void loadQuotationPage("", "reset", statuses);
    }
  }, [activeView, currentUser, quotationPage.initialized, quotationPage.loading, quotationPageKey]);

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

  useEffect(() => {
    if (entryMode !== "lineItems" || !selectedQuotationId || !currentUser) {
      return;
    }

    let cancelled = false;

    async function loadQuotationLineItems() {
      try {
        const response = await apiFetch(`${apiUrl}/api/portal/quotations/${selectedQuotationId}/line-items`);
        const payload = (await response.json()) as { message?: string; items?: PortalLineItemResponse[] };

        if (!response.ok) {
          throw new Error(payload.message || "Failed to load quotation line items");
        }

        if (cancelled) {
          return;
        }

        setLineItemRows(
          payload.items && payload.items.length
            ? payload.items.map(mapPortalLineItemToDraftRow)
            : [createLineItemRow()]
        );
      } catch (loadError) {
        if (!cancelled) {
          setLineItemRows([createLineItemRow()]);
        }
      }
    }

    void loadQuotationLineItems();

    return () => {
      cancelled = true;
    };
  }, [entryMode, selectedQuotationId, currentUser]);

  useEffect(() => {
    if (entryMode !== "order" || !currentUser) {
      return;
    }

    let cancelled = false;

    async function loadOrderItems() {
      try {
        if (editingOrderId) {
          const response = await apiFetch(`${apiUrl}/api/portal/orders/${editingOrderId}/line-items`);
          const payload = (await response.json()) as { message?: string; items?: PortalOrderLineItemResponse[] };
          if (!response.ok) {
            throw new Error(payload.message || "Failed to load order line items");
          }
          if (!cancelled) {
            setOrderLineItems(
              payload.items && payload.items.length
                ? payload.items.map(mapPortalOrderLineItemToDraftRow)
                : [createOrderLineItemRow()]
            );
          }
          return;
        }

        if (!orderForm.quotationId) {
          if (!cancelled) {
            setOrderLineItems([createOrderLineItemRow()]);
          }
          return;
        }

        const response = await apiFetch(`${apiUrl}/api/portal/quotations/${orderForm.quotationId}/line-items`);
        const payload = (await response.json()) as { message?: string; items?: PortalLineItemResponse[] };
        if (!response.ok) {
          throw new Error(payload.message || "Failed to load quotation line items");
        }
        const mappedRows =
          payload.items && payload.items.length
            ? payload.items.map((item) => {
                const product = operations?.products.find((entry) => entry.id === item.productId);
                const calculated = calculateLineItemAmounts(
                  product,
                  item.qty,
                  item.rate,
                  item.transport,
                  item.gstPercent
                );
                return {
                  id: item.id || `${Date.now()}-${Math.random().toString(16).slice(2)}`,
                  productId: item.productId,
                  description:
                    [product?.model, product?.narration].filter(Boolean).join(" - ") ||
                    [product?.name, product?.narration].filter(Boolean).join(" - ") ||
                    product?.name ||
                    product?.productKey ||
                    "Order item",
                  qty: String(item.qty || 1),
                  ratePerUnit: formatAmountInput(item.rate || 0),
                  packingFreight: formatAmountInput(item.transport || 0),
                  unitValue: formatAmountInput(calculated.unitValue),
                  gst18: formatAmountInput(item.gstPercent || 0),
                  totalAmount: formatAmountInput(calculated.totalAmount)
                } satisfies OrderLineItemRow;
              })
            : [createOrderLineItemRow()];

        if (!cancelled) {
          setOrderLineItems(mappedRows);
        }
      } catch {
        if (!cancelled) {
          setOrderLineItems([createOrderLineItemRow()]);
        }
      }
    }

    void loadOrderItems();
    return () => {
      cancelled = true;
    };
  }, [entryMode, editingOrderId, orderForm.quotationId, currentUser, operations?.products]);

  useEffect(() => {
    if (entryMode !== "order") {
      return;
    }

    const total = orderLineItems.reduce((sum, item) => sum + Number(item.totalAmount || 0), 0);
    setOrderForm((current) => ({
      ...current,
      totalAmount: formatAmountInput(total)
    }));
  }, [entryMode, orderLineItems]);

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

  function applyDraftStateToQuotation(input: {
    quotationId: string;
    quotationStatus?: string;
    draftFileUrl?: string;
    driveFolderUrl?: string;
  }) {
    setOperations((current) => {
      if (!current) {
        return current;
      }

      return {
        ...current,
        quotations: current.quotations.map((quotation) =>
          quotation.id === input.quotationId
            ? {
                ...quotation,
                status: input.quotationStatus || quotation.status,
                draftFileUrl: input.draftFileUrl || quotation.draftFileUrl,
                driveFolderUrl: input.driveFolderUrl || quotation.driveFolderUrl,
                draftCreatedTime:
                  input.draftFileUrl && !quotation.draftCreatedTime
                    ? new Date().toISOString()
                    : quotation.draftCreatedTime
              }
            : quotation
        )
      };
    });
  }

  async function handleCreateCustomer(enquiryId: string) {
    const actionKey = `create-customer-${enquiryId}`;
    try {
      setActionState({
        key: actionKey,
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
        key: actionKey,
        label: "Create Customer",
        status: "success",
        message: "Customer linked successfully."
      });
    } catch (actionError) {
      const message =
        actionError instanceof Error ? actionError.message : "Failed to create customer";
      setActionState({
        key: actionKey,
        label: "Create Customer",
        status: "error",
        message
      });
    }
  }

  async function handleGenerateDraftFromEnquiry(enquiryId: string) {
    try {
      setActionState({
        key: `enquiry-draft-${enquiryId}`,
        label: "Generate Draft",
        status: "loading",
        message: "Checking customer, quotation, Drive folder, and draft file..."
      });

      const response = await apiFetch(`${apiUrl}/api/actions/enquiries/${enquiryId}/generate-draft`, {
        method: "POST"
      });
      const payload = (await response.json()) as {
        message?: string;
        quotationNumber?: string;
        draftFileUrl?: string;
        driveFolderUrl?: string;
        quotationId?: string;
        generated?: boolean;
      };

      if (!response.ok) {
        throw new Error(payload.message || "Failed to generate draft");
      }

      try {
        await refreshOperations(false);
      } catch {
        // Keep success state when the follow-up refresh briefly fails.
      }
      setActiveView("quotationDrafts");
      setActionState({
        key: `enquiry-draft-${enquiryId}`,
        label: "Generate Draft",
        status: payload.generated === false ? "error" : "success",
        message:
          payload.message ||
          (payload.generated === false
            ? payload.quotationNumber
              ? `${payload.quotationNumber} is ready. Add line items to generate the draft.`
              : "Quotation shell is ready. Add line items to generate the draft."
            : payload.quotationNumber
              ? `Draft generated successfully for ${payload.quotationNumber}.`
              : "Draft generated successfully.")
      });
    } catch (actionError) {
      const rawMessage = actionError instanceof Error ? actionError.message : "Failed to generate draft";
      const message = rawMessage.toLowerCase().includes("line item")
        ? "Draft could not be generated because no line items exist yet. Add line items first, then try again."
        : rawMessage;
      setActionState({
        key: `enquiry-draft-${enquiryId}`,
        label: "Generate Draft",
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

      const payload = (await response.json()) as {
        message?: string;
        finalPdfUrl?: string;
        previewUrl?: string;
      };

      if (!response.ok) {
        throw new Error(payload.message || "Failed to generate final PDF");
      }

      try {
        await refreshOperations(false);
      } catch {
        // Keep the PDF success state when the follow-up refresh briefly fails.
      }
      setActiveView("approvedQuotations");
      setActionState({
        key: `pdf-generate-${quotationId}`,
        label: "Generate Final PDF",
        status: "success",
        message: payload.finalPdfUrl
          ? "Final PDF generated successfully from the live draft sheet."
          : "Final PDF generated successfully."
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

      const requestController = new AbortController();
      const requestTimeout = window.setTimeout(() => requestController.abort(), 20000);
      let response: Response;

      try {
        response = await apiFetch(`${apiUrl}/api/actions/quotations/${quotationId}/send-${channel}`, {
          method: "POST",
          signal: requestController.signal
        });
      } catch (requestError) {
        if (requestError instanceof DOMException && requestError.name === "AbortError") {
          throw new Error(
            channel === "email"
              ? "Email sending is taking too long. Please check SMTP connectivity and try again."
              : "WhatsApp sending is taking too long. Please try again."
          );
        }

        throw requestError;
      } finally {
        window.clearTimeout(requestTimeout);
      }

      if (!response.ok) {
        const payload = (await response.json()) as { message?: string };
        throw new Error(payload.message || `Failed to ${label.toLowerCase()}`);
      }

      await refreshOperations(false);
      setActiveView(channel === "email" ? "sentQuotations" : "approvedQuotations");
      setActionState({
        key: actionKey,
        label,
        status: "success",
        message:
          channel === "email"
            ? "Quotation emailed successfully."
            : "WhatsApp send request accepted. Waiting for delivery confirmation."
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

      const payload = (await response.json()) as {
        message?: string;
        draftFileUrl?: string;
        driveFolderUrl?: string;
      };

      if (!response.ok) {
        throw new Error(payload.message || "Failed to regenerate quotation draft");
      }

      applyDraftStateToQuotation({
        quotationId,
        quotationStatus: "Draft Quote",
        draftFileUrl: payload.draftFileUrl,
        driveFolderUrl: payload.driveFolderUrl
      });
      try {
        await refreshOperations(false);
      } catch {
        // Keep success state when the follow-up refresh briefly fails.
      }
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
    setProductSearchTerm("");
    setProductDropdownOpen(false);
  }

  async function loadCustomerPage(offset: string, direction: "next" | "previous" | "reset") {
    if (!currentUser) {
      return;
    }

    const previousState = customerPage;
    try {
      setCustomerPage((current) => ({ ...current, loading: true, error: "" }));
      const params = new URLSearchParams({ pageSize: "25" });
      params.set("includeTotal", direction === "reset" ? "1" : "0");
      if (offset) {
        params.set("offset", offset);
      }

      const response = await apiFetch(`${apiUrl}/api/operations/customers?${params.toString()}`);
      if (!response.ok) {
        throw new Error(`Customers API returned ${response.status}`);
      }

      const data = (await response.json()) as CustomersPageResponse;
      setCustomerPage((current) => {
        if (direction === "next") {
          return {
            ...current,
            currentOffset: offset,
            error: "",
            initialized: true,
            loading: false,
            nextOffset: data.nextOffset,
            page: current.page + 1,
            previousOffsets: [...current.previousOffsets, current.currentOffset],
            rows: data.customers,
            totalCount: data.totalCount || current.totalCount
          };
        }

        if (direction === "previous") {
          return {
            ...current,
            currentOffset: offset,
            error: "",
            initialized: true,
            loading: false,
            nextOffset: data.nextOffset,
            page: Math.max(1, current.page - 1),
            previousOffsets: current.previousOffsets.slice(0, -1),
            rows: data.customers,
            totalCount: data.totalCount || current.totalCount
          };
        }

        return {
          ...current,
          currentOffset: "",
          error: "",
          initialized: true,
          loading: false,
          nextOffset: data.nextOffset,
          page: 1,
          previousOffsets: [],
          rows: data.customers,
          totalCount: data.totalCount || current.totalCount
        };
      });
    } catch (loadError) {
      const message = loadError instanceof Error ? loadError.message : "Failed to load customers";
      setCustomerPage({
        ...previousState,
        error: message,
        initialized: previousState.initialized,
        loading: false
      });
      if (message.includes("sign in again")) {
        setAuthError(message);
      }
    }
  }

  async function loadEnquiryPage(
    offset: string,
    direction: "next" | "previous" | "reset",
    statusOverride = enquiryStatusFilter
  ) {
    if (!currentUser) {
      return;
    }

    const previousState = enquiryPage;
    try {
      setEnquiryPage((current) => ({ ...current, loading: true, error: "" }));
      const params = new URLSearchParams({ pageSize: "25" });
      params.set("includeTotal", direction === "reset" ? "1" : "0");
      if (offset) {
        params.set("offset", offset);
      }
      if (statusOverride !== "All") {
        params.set("status", statusOverride);
      }

      const response = await apiFetch(`${apiUrl}/api/operations/enquiries?${params.toString()}`);
      if (!response.ok) {
        throw new Error(`Enquiries API returned ${response.status}`);
      }

      const data = (await response.json()) as EnquiriesPageResponse;
      setEnquiryPage((current) => {
        if (direction === "next") {
          return {
            ...current,
            currentOffset: offset,
            error: "",
            initialized: true,
            loading: false,
            nextOffset: data.nextOffset,
            page: current.page + 1,
            previousOffsets: [...current.previousOffsets, current.currentOffset],
            rows: data.enquiries,
            totalCount: data.totalCount || current.totalCount
          };
        }

        if (direction === "previous") {
          return {
            ...current,
            currentOffset: offset,
            error: "",
            initialized: true,
            loading: false,
            nextOffset: data.nextOffset,
            page: Math.max(1, current.page - 1),
            previousOffsets: current.previousOffsets.slice(0, -1),
            rows: data.enquiries,
            totalCount: data.totalCount || current.totalCount
          };
        }

        return {
          ...current,
          currentOffset: "",
          error: "",
          initialized: true,
          loading: false,
          nextOffset: data.nextOffset,
          page: 1,
          previousOffsets: [],
          rows: data.enquiries,
          totalCount: data.totalCount || current.totalCount
        };
      });
    } catch (loadError) {
      const message = loadError instanceof Error ? loadError.message : "Failed to load enquiries";
      setEnquiryPage({
        ...previousState,
        error: message,
        loading: false
      });
      if (message.includes("sign in again")) {
        setAuthError(message);
      }
    }
  }

  async function loadQuotationPage(
    offset: string,
    direction: "next" | "previous" | "reset",
    statuses: string[]
  ) {
    if (!currentUser) {
      return;
    }

    const key = statuses.join("|") || "all";
    const previousState = quotationPage;
    try {
      setQuotationPageKey(key);
      setQuotationPage((current) => ({ ...current, loading: true, error: "" }));
      const params = new URLSearchParams({ pageSize: "25" });
      params.set("includeTotal", direction === "reset" ? "1" : "0");
      if (offset) {
        params.set("offset", offset);
      }
      if (statuses.length) {
        params.set("statuses", statuses.join(","));
      }

      const response = await apiFetch(`${apiUrl}/api/operations/quotations?${params.toString()}`);
      if (!response.ok) {
        throw new Error(`Quotations API returned ${response.status}`);
      }

      const data = (await response.json()) as QuotationsPageResponse;
      setQuotationPage((current) => {
        if (direction === "next") {
          return {
            ...current,
            currentOffset: offset,
            error: "",
            initialized: true,
            loading: false,
            nextOffset: data.nextOffset,
            page: current.page + 1,
            previousOffsets: [...current.previousOffsets, current.currentOffset],
            rows: data.quotations,
            totalCount: data.totalCount || current.totalCount
          };
        }

        if (direction === "previous") {
          return {
            ...current,
            currentOffset: offset,
            error: "",
            initialized: true,
            loading: false,
            nextOffset: data.nextOffset,
            page: Math.max(1, current.page - 1),
            previousOffsets: current.previousOffsets.slice(0, -1),
            rows: data.quotations,
            totalCount: data.totalCount || current.totalCount
          };
        }

        return {
          ...current,
          currentOffset: "",
          error: "",
          initialized: true,
          loading: false,
          nextOffset: data.nextOffset,
          page: 1,
          previousOffsets: [],
          rows: data.quotations,
          totalCount: data.totalCount || current.totalCount
        };
      });
    } catch (loadError) {
      const message = loadError instanceof Error ? loadError.message : "Failed to load quotations";
      setQuotationPage({
        ...previousState,
        error: message,
        loading: false
      });
      if (message.includes("sign in again")) {
        setAuthError(message);
      }
    }
  }

  function openCustomerEntry(customer?: CustomerRecord) {
    setEntryMode("customer");
    setEditingCustomerId(customer?.id || "");
    setCustomerForm(
      customer
        ? {
            customerName: customer.customerName,
            company: customer.company,
            phone: customer.phone,
            email: customer.email,
            address: customer.address,
            state: customer.state,
            city: customer.city,
            pincode: customer.pincode,
            customerType: customer.customerType || "Domestic"
          }
        : createBlankCustomerForm()
    );
  }

  function openQuotationEntry() {
    setEntryMode("quotation");
    setQuotationForm(createBlankQuotationForm());
  }

function openOrderEntry(order?: OrderRecord, quotation?: QuotationRecord) {
    const linkedEnquiry = quotation?.linkedEnquiryId ? enquiryLookup.get(quotation.linkedEnquiryId) : undefined;
    setEntryMode("order");
    setEditingOrderId(order?.id || "");
    setOrderForm(
      order
        ? {
            quotationId: order.linkedQuotationId,
            customerId: order.linkedCustomerId,
            enquiryId: order.linkedEnquiryId || "",
            orderDate: order.orderDate ? order.orderDate.slice(0, 10) : "",
            orderStatus: (order.orderStatus as OrderFormState["orderStatus"]) || "Confirmed",
            totalAmount: formatAmountInput(order.totalAmount || order.orderValue || 0),
            orderNotes: order.orderNotes || "",
            paymentStatus: (order.paymentStatus as OrderFormState["paymentStatus"]) || "Pending",
            address: order.address || "",
            state: order.state || "",
            city: order.city || "",
            pincode: order.pincode || ""
          }
        : {
            quotationId: quotation?.id || "",
            customerId: quotation?.linkedCustomerId || "",
            enquiryId: quotation?.linkedEnquiryId || "",
            orderDate: new Date().toISOString().slice(0, 10),
            orderStatus: "Confirmed",
            totalAmount: formatAmountInput(quotation?.quotationGrandTotal || 0),
            orderNotes: "",
            paymentStatus: "Pending",
            address: linkedEnquiry?.destinationAddress || linkedEnquiry?.address || "",
            state: linkedEnquiry?.destinationState || linkedEnquiry?.state || "",
            city: linkedEnquiry?.destinationCity || linkedEnquiry?.city || "",
            pincode: linkedEnquiry?.destinationPincode || linkedEnquiry?.pincode || ""
          }
    );
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
      address: enquiry.address,
      state: enquiry.state,
      city: enquiry.city,
      pincode: enquiry.pincode,
      destinationAddress: enquiry.destinationAddress,
      destinationState: enquiry.destinationState,
      destinationCity: enquiry.destinationCity,
      destinationPincode: enquiry.destinationPincode,
      requirementSummary: enquiry.requirementSummary,
      potentialProduct: enquiry.potentialProduct
    });
    setEnquiryFieldErrors({});
    setDestinationSameAsMain(false);
    setCustomerSearchTerm(
      customer ? `${customer.clientId || "No ID"} - ${customer.customerName || "Unnamed customer"}` : ""
    );
    setCustomerDropdownOpen(false);
    const matchedProduct = (operations?.products || []).find((product) => product.id === enquiry.potentialProduct);
    setProductSearchTerm(
      matchedProduct
        ? matchedProduct.name || matchedProduct.model || matchedProduct.productKey
        : ""
    );
    setProductDropdownOpen(false);
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
    setEditingOrderId("");
    setEnquiryFieldErrors({});
    setDestinationSameAsMain(false);
    setCustomerSearchTerm("");
    setCustomerDropdownOpen(false);
    setProductSearchTerm("");
    setProductDropdownOpen(false);
    setOrderForm(createBlankOrderForm());
    setOrderLineItems([createOrderLineItemRow()]);
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

  const orderByQuotationId = useMemo(() => {
    return new Map((operations?.orders ?? []).map((order) => [order.linkedQuotationId, order]));
  }, [operations?.orders]);

  const availableQuotationOptions = useMemo(() => {
    return [...(operations?.quotations ?? [])].sort((left, right) => {
      const leftNumber = left.quotationNumber || "";
      const rightNumber = right.quotationNumber || "";
      return rightNumber.localeCompare(leftNumber);
    });
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

  const filteredProducts = useMemo(() => {
    const search = productSearchTerm.trim().toLowerCase();
    const products = operations?.products || [];

    if (!search) {
      return products;
    }

    return products.filter((product) =>
      `${product.productKey} ${product.model} ${product.name} ${product.narration}`
        .toLowerCase()
        .includes(search)
    );
  }, [operations?.products, productSearchTerm]);

  const enquiryStatusOptions = useMemo(() => {
    const uniqueStatuses = Array.from(
      new Set([
        ...enquiryWorkflowStatuses,
        ...(operations?.enquiries || []).map((enquiry) => enquiry.parserStatus || "")
      ])
    ).filter(Boolean);
    return ["All", ...uniqueStatuses];
  }, [operations?.enquiries]);

  const filteredEnquiries = useMemo(() => {
    return enquiryPage.initialized ? enquiryPage.rows : operations?.enquiries || [];
  }, [enquiryPage.initialized, enquiryPage.rows, operations?.enquiries]);

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
      state: customer?.state || current.state,
      city: customer?.city || current.city,
      pincode: customer?.pincode || current.pincode
    }));
  }

  function handleProductAutofill(productId: string) {
    const product = (operations?.products || []).find((item) => item.id === productId);
    setProductSearchTerm(product ? product.name || product.model || product.productKey : "");
    setProductDropdownOpen(false);
    clearEnquiryFieldError("potentialProduct");
    setEnquiryForm((current) => ({
      ...current,
      potentialProduct: productId,
      requirementSummary: product
        ? [product.model, product.narration].filter(Boolean).join(" - ") ||
          [product.name, product.narration].filter(Boolean).join(" - ") ||
          product.model ||
          product.name ||
          product.productKey
        : current.requirementSummary
    }));
  }

  async function handleSubmitEnquiry() {
    if (enquirySubmitInFlightRef.current) {
      return;
    }

    const isEditing = Boolean(editingEnquiryId);
    const formActionLabel = isEditing ? "Update Enquiry" : "Create Enquiry";
    const normalizedPhone = normalizePhoneInput(enquiryForm.phone);
    const normalizedEmail = enquiryForm.email.trim();
    const nextFieldErrors: EnquiryFieldErrors = {};

    if (!enquiryForm.leadName.trim()) {
      nextFieldErrors.leadName = "Lead name is required.";
    }

    if (enquiryForm.phone.trim() && !isValidTenDigitPhone(enquiryForm.phone)) {
      nextFieldErrors.phone = "Phone number must be exactly 10 digits.";
    }

    if (normalizedEmail && !isValidEmail(normalizedEmail)) {
      nextFieldErrors.email = "Email must include @.";
    }

    if (!enquiryForm.state.trim()) {
      nextFieldErrors.state = "State is required.";
    }

    if (enquiryForm.pincode.trim() && !/^\d{6}$/.test(normalizePincodeInput(enquiryForm.pincode))) {
      nextFieldErrors.pincode = "Main pincode must be a valid 6-digit number.";
    }

    if (!enquiryForm.potentialProduct) {
      nextFieldErrors.potentialProduct = "Product is required.";
    }

    if (Object.keys(nextFieldErrors).length) {
      setEnquiryFieldErrors(nextFieldErrors);
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "error",
        message: firstEnquiryFieldError(nextFieldErrors)
      });
      return;
    }

    setEnquiryFieldErrors({});

    try {
      enquirySubmitInFlightRef.current = true;
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
            address: enquiryForm.address,
            state: enquiryForm.state,
            city: enquiryForm.city,
            pincode: normalizePincodeInput(enquiryForm.pincode),
            destinationAddress: enquiryForm.destinationAddress,
            destinationState: enquiryForm.destinationState,
            destinationCity: enquiryForm.destinationCity,
            destinationPincode: normalizePincodeInput(enquiryForm.destinationPincode),
            requirementSummary: enquiryForm.requirementSummary,
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
      const backendFieldErrors: EnquiryFieldErrors = {};

      if (message.toLowerCase().includes("lead name")) {
        backendFieldErrors.leadName = "Lead name is required.";
      }
      if (message.toLowerCase().includes("phone number")) {
        backendFieldErrors.phone = "Phone number must be exactly 10 digits.";
      }
      if (message.toLowerCase().includes("email")) {
        backendFieldErrors.email = "Enter a valid email address.";
      }
      if (message.toLowerCase().includes("select a product")) {
        backendFieldErrors.potentialProduct = "Product is required.";
      }
      if (message.toLowerCase().includes("pincode")) {
        backendFieldErrors.pincode = "Main pincode must be a valid 6-digit number.";
      }

      setEnquiryFieldErrors(backendFieldErrors);
      setActionState({
        key: "portal-enquiry",
        label: formActionLabel,
        status: "error",
        message
      });
    } finally {
      enquirySubmitInFlightRef.current = false;
    }
  }

  async function handleSubmitCustomer() {
    if (customerSubmitInFlightRef.current) {
      return;
    }

    const isEditing = Boolean(editingCustomerId);
    const actionKey = "portal-customer";
    const label = isEditing ? "Update Customer" : "Create Customer";

    try {
      customerSubmitInFlightRef.current = true;
      setActionState({
        key: actionKey,
        label,
        status: "loading",
        message: isEditing ? "Updating customer..." : "Creating customer..."
      });

      const response = await apiFetch(
        isEditing ? `${apiUrl}/api/portal/customers/${editingCustomerId}` : `${apiUrl}/api/portal/customers`,
        {
          method: isEditing ? "PATCH" : "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            customerName: customerForm.customerName,
            company: customerForm.company,
            phone: normalizePhoneInput(customerForm.phone),
            email: customerForm.email.trim(),
            address: customerForm.address,
            state: customerForm.state,
            city: customerForm.city,
            pincode: normalizePincodeInput(customerForm.pincode),
            customerType: customerForm.customerType
          })
        }
      );
      const payload = (await response.json()) as { message?: string };
      if (!response.ok) {
        throw new Error(payload.message || "Failed to save customer");
      }

      setActionState({
        key: actionKey,
        label,
        status: "success",
        message: isEditing ? "Customer updated successfully." : "Customer created successfully."
      });
      setCustomerForm(createBlankCustomerForm());
      closeEntryPanel();
      setActiveView("customers");
      setCustomerPage((current) => ({ ...current, initialized: false }));
      void refreshOperations(true);
      } catch (error) {
        setActionState({
          key: actionKey,
          label,
          status: "error",
          message: error instanceof Error ? error.message : "Failed to save customer"
        });
      } finally {
        customerSubmitInFlightRef.current = false;
      }
    }

  async function handleSubmitQuotation() {
    if (quotationSubmitInFlightRef.current) {
      return;
    }

    const actionKey = "portal-quotation";
    try {
      quotationSubmitInFlightRef.current = true;
      setActionState({
        key: actionKey,
        label: "Create Quotation",
        status: "loading",
        message: "Creating quotation shell from enquiry..."
      });

      const response = await apiFetch(`${apiUrl}/api/portal/quotations`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ enquiryId: quotationForm.enquiryId })
      });
      const payload = (await response.json()) as {
        message?: string;
        quotation?: { id: string; fields?: { ["Quotation Number"]?: string } };
      };
      if (!response.ok) {
        throw new Error(payload.message || "Failed to create quotation");
      }

      setActionState({
        key: actionKey,
        label: "Create Quotation",
        status: "success",
        message: "Quotation created. Add line items to generate the draft."
      });
      const nextQuotationId = payload.quotation?.id || "";
      closeEntryPanel();
      void refreshOperations(true);
      if (nextQuotationId) {
        openLineItemEntry(nextQuotationId);
      }
      } catch (error) {
        setActionState({
          key: actionKey,
          label: "Create Quotation",
          status: "error",
          message: error instanceof Error ? error.message : "Failed to create quotation"
        });
      } finally {
        quotationSubmitInFlightRef.current = false;
      }
    }

  async function handleSubmitOrder() {
    if (orderSubmitInFlightRef.current) {
      return;
    }

    const isEditing = Boolean(editingOrderId);
    const actionKey = isEditing ? `order-update-${editingOrderId}` : `order-create-${orderForm.quotationId}`;
    const url = isEditing
      ? `${apiUrl}/api/portal/orders/${editingOrderId}`
      : `${apiUrl}/api/actions/quotations/${orderForm.quotationId}/create-order`;
    const method = isEditing ? "PATCH" : "POST";
    const normalizedOrderItems = orderLineItems.map((item) => {
      const qty = Math.max(1, Number(item.qty || 0));
      const ratePerUnit = Number(item.ratePerUnit || 0);
      const packingFreight = Number(item.packingFreight || 0);
      const gst18 = Number(item.gst18 || 0);
      const unitValue = Number((qty * ratePerUnit).toFixed(2));
      const totalAmount = Number((unitValue + packingFreight * qty + gst18 * qty).toFixed(2));

      return {
        id: item.id,
        productId: item.productId,
        description: item.description.trim(),
        qty,
        ratePerUnit,
        packingFreight,
        unitValue,
        gst18,
        totalAmount
      };
    });

    if (normalizedOrderItems.some((item) => !item.description)) {
      setActionState({
        key: actionKey,
        label: isEditing ? "Update Order" : "Create Order",
        status: "error",
        message: "Each order line item needs a description before saving."
      });
      return;
    }

      try {
        orderSubmitInFlightRef.current = true;
        setActionState({
          key: actionKey,
          label: isEditing ? "Update Order" : "Create Order",
        status: "loading",
        message: isEditing ? "Saving order changes..." : "Creating order from quotation..."
      });

      const response = await apiFetch(url, {
        method,
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            quotationId: orderForm.quotationId,
            customerId: orderForm.customerId,
            enquiryId: orderForm.enquiryId,
            orderDate: orderForm.orderDate,
            orderStatus: orderForm.orderStatus,
            totalAmount: Number(orderForm.totalAmount || 0),
            orderNotes: orderForm.orderNotes,
            paymentStatus: orderForm.paymentStatus,
            address: orderForm.address,
            state: orderForm.state,
            city: orderForm.city,
            pincode: normalizePincodeInput(orderForm.pincode),
            items: normalizedOrderItems
          })
        });
      const payload = (await response.json()) as { message?: string };
      if (!response.ok) {
        throw new Error(payload.message || "Failed to save order");
      }

      setActionState({
        key: actionKey,
        label: isEditing ? "Update Order" : "Create Order",
        status: "success",
        message: isEditing ? "Order updated successfully." : "Order created successfully."
      });
      setOrderLineItems([createOrderLineItemRow()]);
      closeEntryPanel();
      setActiveView("orders");
      void refreshOperations(true);
      } catch (error) {
        setActionState({
          key: actionKey,
          label: isEditing ? "Update Order" : "Create Order",
          status: "error",
          message: error instanceof Error ? error.message : "Failed to save order"
        });
      } finally {
        orderSubmitInFlightRef.current = false;
      }
    }

  async function handleMarkQuotationSent(quotationId: string) {
    try {
      setActionState({
        key: `quotation-mark-sent-${quotationId}`,
        label: "Mark as Sent",
        status: "loading",
        message: "Moving quotation to sent stage..."
      });
      const response = await apiFetch(`${apiUrl}/api/actions/quotations/${quotationId}/mark-sent`, {
        method: "POST"
      });
      const payload = (await response.json()) as { message?: string };
      if (!response.ok) {
        throw new Error(payload.message || "Failed to mark quotation as sent");
      }

      try {
        await refreshOperations(false);
      } catch {
        // Keep success state when the follow-up refresh briefly fails.
      }
      setActiveView("sentQuotations");
      setActionState({
        key: `quotation-mark-sent-${quotationId}`,
        label: "Mark as Sent",
        status: "success",
        message: "Quotation moved to Sent Quotations."
      });
    } catch (error) {
      setActionState({
        key: `quotation-mark-sent-${quotationId}`,
        label: "Mark as Sent",
        status: "error",
        message: error instanceof Error ? error.message : "Failed to mark quotation as sent"
      });
    }
  }

function updateLineItemRow(
  rowId: string,
  field: "productId" | "qty" | "rate" | "transport" | "gstPercent",
  value: string
) {
    setLineItemRows((current) =>
      current.map((row) => {
        if (row.id !== rowId) {
          return row;
        }

        const nextValue =
          field === "rate" || field === "transport" || field === "gstPercent"
            ? normalizeDecimalInput(value)
            : value;
        const nextRow = { ...row, [field]: nextValue };

        const product = operations?.products.find((item) => item.id === nextRow.productId);
        if (field === "productId" && product) {
          nextRow.rate = formatAmountInput(Number(product.bulkSalePrice || product.mrp || 0));
          nextRow.transport = formatAmountInput(Number(product.transportCharge || 0));
          nextRow.gstPercent = formatAmountInput(Number(product.gstPercent || 0));
        }

        return nextRow;
      })
    );
  }

  function updateOrderLineItemRow(
    rowId: string,
    field: keyof OrderLineItemRow,
    value: string
  ) {
    setOrderLineItems((current) =>
      current.map((row) => {
        if (row.id !== rowId) {
          return row;
        }

        const nextRow = {
          ...row,
          [field]:
            field === "qty"
              ? value.replace(/[^\d]/g, "")
              : field === "ratePerUnit" || field === "packingFreight" || field === "gst18"
                ? normalizeDecimalInput(value)
                : value
        };

        const qty = Math.max(1, Number(nextRow.qty || 0));
        const ratePerUnit = Number(nextRow.ratePerUnit || 0);
        const packingFreight = Number(nextRow.packingFreight || 0);
        const gst18 = Number(nextRow.gst18 || 0);
        const unitValue = Number((qty * ratePerUnit).toFixed(2));
        const totalAmount = Number((unitValue + packingFreight * qty + gst18 * qty).toFixed(2));

        nextRow.unitValue = formatAmountInput(unitValue);
        nextRow.totalAmount = formatAmountInput(totalAmount);

        return nextRow;
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

  function addOrderLineItemRow() {
    setOrderLineItems((current) => [...current, createOrderLineItemRow()]);
  }

  function removeOrderLineItemRow(rowId: string) {
    setOrderLineItems((current) => (current.length > 1 ? current.filter((row) => row.id !== rowId) : current));
  }

  async function handleSubmitLineItems() {
    if (lineItemsSubmitInFlightRef.current) {
      return;
    }

    if (lineItemRows.some((row) => !row.productId)) {
      setActionState({
        key: "portal-line-items",
        label: "Create Line Items",
        status: "error",
        message: "Select a product for every line item before creating them."
      });
      return;
    }

    if (lineItemRows.some((row) => !isValidPositiveAmount(row.rate))) {
      setActionState({
        key: "portal-line-items",
        label: "Create Line Items",
        status: "error",
        message: "Rate must be a number greater than zero. Decimals are allowed."
      });
      return;
    }

      try {
        lineItemsSubmitInFlightRef.current = true;
        setActionState({
          key: "portal-line-items",
          label: "Create Line Items",
        status: "loading",
        message: "Creating quotation line items..."
      });

      const response = await apiFetch(`${apiUrl}/api/portal/quotation-line-items`, {
        method: "PUT",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          quotationId: selectedQuotationId,
          items: lineItemRows.map((row) => ({
            productId: row.productId,
            qty: row.qty,
            rate: row.rate,
            transport: row.transport || "0",
            gstPercent: row.gstPercent || "0"
          }))
        })
      });

      const payload = (await response.json()) as {
        message?: string;
        quotationId?: string;
        quotationNumber?: string;
        createdCount?: number;
        quotationStatus?: string;
        draftFileUrl?: string;
        driveFolderUrl?: string;
      };

      if (!response.ok) {
        throw new Error(payload.message || "Failed to create line items");
      }

      if (payload.quotationId) {
        applyDraftStateToQuotation({
          quotationId: payload.quotationId,
          quotationStatus: payload.quotationStatus,
          draftFileUrl: payload.draftFileUrl,
          driveFolderUrl: payload.driveFolderUrl
        });
      }

      setActionState({
        key: "portal-line-items",
        label: "Create Line Items",
        status: "success",
        message: payload.draftFileUrl
          ? `${payload.createdCount || lineItemRows.length} line items saved for ${payload.quotationNumber || "quotation"}. Draft updated automatically.`
          : `${payload.createdCount || lineItemRows.length} line items saved for ${payload.quotationNumber || "quotation"}. If the draft does not appear, use Generate Draft from the quotation row.`
      });
      setLineItemRows([createLineItemRow()]);
      closeEntryPanel();
      setActiveView("quotationDrafts");
      try {
        await refreshOperations(false);
      } catch {
        // Keep success state when the follow-up refresh briefly fails.
      }
      } catch (submitError) {
        const message =
          submitError instanceof Error ? submitError.message : "Failed to create line items";
        setActionState({
          key: "portal-line-items",
          label: "Create Line Items",
          status: "error",
          message
        });
      } finally {
        lineItemsSubmitInFlightRef.current = false;
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
      !isValidPositiveAmount(row.rate) ||
      (row.transport.trim() !== "" && Number(row.transport) < 0) ||
      (row.gstPercent.trim() !== "" && Number(row.gstPercent) < 0)
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
                onChange={(event) => {
                  clearEnquiryFieldError("leadName");
                  setEnquiryForm((current) => ({ ...current, leadName: event.target.value }));
                }}
              />
              {enquiryFieldErrors.leadName ? <span className="field-error">{enquiryFieldErrors.leadName}</span> : null}
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
              <span>Pincode</span>
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
              <select
                value={enquiryForm.state}
                onChange={(event) => {
                  clearEnquiryFieldError("state");
                  setEnquiryForm((current) => ({ ...current, state: event.target.value }));
                }}
              >
                <option value="">Select state</option>
                {stateOptions.map((stateOption) => (
                  <option key={stateOption} value={stateOption}>
                    {stateOption}
                  </option>
                ))}
              </select>
              {enquiryFieldErrors.state ? <span className="field-error">{enquiryFieldErrors.state}</span> : null}
            </label>
            <label>
              <span>City</span>
              <input
                value={enquiryForm.city}
                onChange={(event) => {
                  clearEnquiryFieldError("city");
                  setEnquiryForm((current) => ({ ...current, city: event.target.value }));
                }}
              />
              {enquiryFieldErrors.city ? <span className="field-error">{enquiryFieldErrors.city}</span> : null}
            </label>
            <label>
              <span>Product *</span>
              <div className="searchable-dropdown">
                <div className="searchable-dropdown-input">
                  <input
                    value={productSearchTerm}
                    placeholder="Search by product name, model, or key"
                    onFocus={() => setProductDropdownOpen(true)}
                    onClick={() => setProductDropdownOpen(true)}
                    onBlur={() => {
                      window.setTimeout(() => setProductDropdownOpen(false), 120);
                    }}
                    onChange={(event) => {
                      const value = event.target.value;
                      clearEnquiryFieldError("potentialProduct");
                      setProductSearchTerm(value);
                      setProductDropdownOpen(true);
                      if (!value.trim()) {
                        setEnquiryForm((current) => ({
                          ...current,
                          potentialProduct: ""
                        }));
                      }
                    }}
                  />
                  <button
                    type="button"
                    className="search-dropdown-toggle"
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => {
                      setProductDropdownOpen((current) => !current);
                    }}
                  >
                    {productDropdownOpen ? "^" : "v"}
                  </button>
                </div>
                <span className="search-dropdown-hint">
                  Search and choose the product to map this enquiry.
                </span>
                {productDropdownOpen ? (
                  <div className="search-dropdown-list">
                    {filteredProducts.length ? (
                      filteredProducts.map((productOption) => (
                        <button
                          key={productOption.id}
                          className="search-dropdown-option"
                          type="button"
                          onMouseDown={(event) => {
                            event.preventDefault();
                            handleProductAutofill(productOption.id);
                          }}
                        >
                          <strong>{productOption.name || productOption.model || productOption.productKey}</strong>
                          <span>
                            {[productOption.model, productOption.productKey, productOption.narration]
                              .filter(Boolean)
                              .join(" | ") || "Product"}
                          </span>
                        </button>
                      ))
                    ) : (
                      <div className="search-dropdown-empty">No matching products found.</div>
                    )}
                  </div>
                ) : null}
              </div>
              {enquiryFieldErrors.potentialProduct ? (
                <span className="field-error">{enquiryFieldErrors.potentialProduct}</span>
              ) : null}
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

    if (entryMode === "customer") {
      const isSavingCustomer = actionState?.key === "portal-customer" && actionState.status === "loading";

      return (
        <section className="entry-overlay">
          <div className="entry-backdrop" onClick={closeEntryPanel} />
          <section className="panel entry-panel entry-modal" ref={entryModalRef}>
            <div className="panel-header panel-header-tight">
              <div>
                <p className="eyebrow">Customer</p>
                <h2>{editingCustomerId ? "Edit customer" : "Create customer"}</h2>
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
              <label>
                <span>Customer name</span>
                <input value={customerForm.customerName} onChange={(event) => setCustomerForm((current) => ({ ...current, customerName: event.target.value }))} />
              </label>
              <label>
                <span>Company</span>
                <input value={customerForm.company} onChange={(event) => setCustomerForm((current) => ({ ...current, company: event.target.value }))} />
              </label>
              <label>
                <span>Phone</span>
                <input value={customerForm.phone} onChange={(event) => setCustomerForm((current) => ({ ...current, phone: normalizePhoneInput(event.target.value) }))} />
              </label>
              <label>
                <span>Email</span>
                <input type="email" value={customerForm.email} onChange={(event) => setCustomerForm((current) => ({ ...current, email: event.target.value }))} />
              </label>
              <label className="form-span-2">
                <span>Address</span>
                <textarea rows={2} value={customerForm.address} onChange={(event) => setCustomerForm((current) => ({ ...current, address: event.target.value }))} />
              </label>
              <label>
                <span>State</span>
                <select value={customerForm.state} onChange={(event) => setCustomerForm((current) => ({ ...current, state: event.target.value }))}>
                  <option value="">Select state</option>
                  {stateOptions.map((stateOption) => (
                    <option key={stateOption} value={stateOption}>
                      {stateOption}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                <span>City</span>
                <input value={customerForm.city} onChange={(event) => setCustomerForm((current) => ({ ...current, city: event.target.value }))} />
              </label>
              <label>
                <span>Pincode</span>
                <input value={customerForm.pincode} onChange={(event) => setCustomerForm((current) => ({ ...current, pincode: normalizePincodeInput(event.target.value) }))} />
              </label>
              <label>
                <span>Customer type</span>
                <select value={customerForm.customerType} onChange={(event) => setCustomerForm((current) => ({ ...current, customerType: event.target.value }))}>
                  <option value="Domestic">Domestic</option>
                  <option value="Export">Export</option>
                </select>
              </label>
            </div>
            <div className="entry-actions">
              <button className="action-inline-button" type="button" onClick={() => void handleSubmitCustomer()} disabled={isSavingCustomer}>
                {isSavingCustomer ? "Saving..." : editingCustomerId ? "Update customer" : "Create customer"}
              </button>
            </div>
          </section>
        </section>
      );
    }

    if (entryMode === "quotation") {
      const isSavingQuotation = actionState?.key === "portal-quotation" && actionState.status === "loading";

      return (
        <section className="entry-overlay">
          <div className="entry-backdrop" onClick={closeEntryPanel} />
          <section className="panel entry-panel entry-modal" ref={entryModalRef}>
            <div className="panel-header panel-header-tight">
              <div>
                <p className="eyebrow">Quotation</p>
                <h2>Create quotation from enquiry</h2>
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
                <span>Enquiry</span>
                <select value={quotationForm.enquiryId} onChange={(event) => setQuotationForm((current) => ({ ...current, enquiryId: event.target.value }))}>
                  <option value="">Select enquiry</option>
                  {(operations?.enquiries || []).map((enquiry) => (
                    <option key={enquiry.id} value={enquiry.id}>
                      {enquiry.enquiryId} - {enquiry.leadName || enquiry.phone || "Enquiry"}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            <div className="entry-actions">
              <button className="action-inline-button" type="button" onClick={() => void handleSubmitQuotation()} disabled={isSavingQuotation}>
                {isSavingQuotation ? "Creating..." : "Create quotation"}
              </button>
            </div>
          </section>
        </section>
      );
    }

    if (entryMode === "order") {
      const isSavingOrder =
        actionState?.key === `order-update-${editingOrderId}` ||
        actionState?.key === `order-create-${orderForm.quotationId}`
          ? actionState.status === "loading"
          : false;

      return (
        <section className="entry-overlay">
          <div className="entry-backdrop" onClick={closeEntryPanel} />
          <section className="panel entry-panel entry-modal" ref={entryModalRef}>
            <div className="panel-header panel-header-tight">
              <div>
                <p className="eyebrow">Order</p>
                <h2>{editingOrderId ? "Edit order" : "Create order from quotation"}</h2>
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
              <label>
                <span>Order date</span>
                <input type="date" value={orderForm.orderDate} onChange={(event) => setOrderForm((current) => ({ ...current, orderDate: event.target.value }))} />
              </label>
              <label>
                <span>Order status</span>
                <select value={orderForm.orderStatus} onChange={(event) => setOrderForm((current) => ({ ...current, orderStatus: event.target.value as OrderFormState["orderStatus"] }))}>
                  {["Confirmed", "Processing", "Shipped", "Delivered", "Cancelled"].map((status) => (
                    <option key={status} value={status}>
                      {status}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                <span>Total amount</span>
                <input type="number" step="0.01" value={orderForm.totalAmount} onChange={(event) => setOrderForm((current) => ({ ...current, totalAmount: normalizeDecimalInput(event.target.value) }))} />
              </label>
              <label>
                <span>Payment status</span>
                <select value={orderForm.paymentStatus} onChange={(event) => setOrderForm((current) => ({ ...current, paymentStatus: event.target.value as OrderFormState["paymentStatus"] }))}>
                  {["Paid", "Pending", "Half Payment"].map((status) => (
                    <option key={status} value={status}>
                      {status}
                    </option>
                  ))}
                </select>
              </label>
              <label className="form-span-2">
                <span>Address</span>
                <textarea rows={2} value={orderForm.address} onChange={(event) => setOrderForm((current) => ({ ...current, address: event.target.value }))} />
              </label>
              <label>
                <span>State</span>
                <select value={orderForm.state} onChange={(event) => setOrderForm((current) => ({ ...current, state: event.target.value }))}>
                  <option value="">Select state</option>
                  {stateOptions.map((status) => (
                    <option key={status} value={status}>
                      {status}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                <span>City</span>
                <input value={orderForm.city} onChange={(event) => setOrderForm((current) => ({ ...current, city: event.target.value }))} />
              </label>
              <label>
                <span>Pincode</span>
                <input inputMode="numeric" value={orderForm.pincode} onChange={(event) => setOrderForm((current) => ({ ...current, pincode: normalizePincodeInput(event.target.value) }))} />
              </label>
              <label className="form-span-2">
                <span>Order notes</span>
                <textarea rows={3} value={orderForm.orderNotes} onChange={(event) => setOrderForm((current) => ({ ...current, orderNotes: event.target.value }))} />
              </label>
            </div>
            <div className="line-item-builder order-line-item-builder">
              {orderLineItems.map((row, index) => (
                <div className="line-item-row" key={row.id}>
                  <label className="line-item-product">
                    <span>S.No. {index + 1}</span>
                    <input value={row.description} onChange={(event) => updateOrderLineItemRow(row.id, "description", event.target.value)} />
                  </label>
                  <label className="line-item-qty">
                    <span>Qty</span>
                    <input type="number" min="1" value={row.qty} onChange={(event) => updateOrderLineItemRow(row.id, "qty", event.target.value)} />
                  </label>
                  <label className="line-item-qty">
                    <span>Rate Per Unit</span>
                    <input type="number" min="0" step="0.01" value={row.ratePerUnit} onChange={(event) => updateOrderLineItemRow(row.id, "ratePerUnit", event.target.value)} />
                  </label>
                  <label className="line-item-qty">
                    <span>Packing & Freight</span>
                    <input type="number" min="0" step="0.01" value={row.packingFreight} onChange={(event) => updateOrderLineItemRow(row.id, "packingFreight", event.target.value)} />
                  </label>
                  <label className="line-item-qty">
                    <span>Unit Value</span>
                    <input readOnly value={row.unitValue} />
                  </label>
                  <label className="line-item-qty">
                    <span>GST 18%</span>
                    <input type="number" min="0" step="0.01" value={row.gst18} onChange={(event) => updateOrderLineItemRow(row.id, "gst18", event.target.value)} />
                  </label>
                  <label className="line-item-total">
                    <span>Total Amount</span>
                    <input readOnly value={row.totalAmount} />
                  </label>
                  <button className="icon-action-button danger" type="button" title="Remove order line item" onClick={() => removeOrderLineItemRow(row.id)}>
                    X
                  </button>
                </div>
              ))}
            </div>
            <div className="entry-actions">
              <button className="action-inline-button neutral" type="button" onClick={addOrderLineItemRow}>
                Add Order Line Item
              </button>
              <button className="action-inline-button" type="button" onClick={() => void handleSubmitOrder()} disabled={isSavingOrder}>
                {isSavingOrder ? "Saving..." : editingOrderId ? "Update order" : "Create order"}
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
              const amounts = calculateLineItemAmounts(
                product,
                row.qty,
                row.rate,
                row.transport,
                row.gstPercent
              );

              return (
                <div className="line-item-row" key={row.id}>
                  <label className="line-item-product">
                    <span>Product {index + 1}</span>
                    {row.existing ? (
                      <div className="locked-field">
                        <strong>{product?.name || product?.model || product?.productKey || "Selected product"}</strong>
                        <span>{[product?.model, product?.productKey].filter(Boolean).join(" | ") || "Product locked after insertion"}</span>
                      </div>
                    ) : (
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
                    )}
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
                  <label className="line-item-qty">
                    <span>Rate</span>
                    <input
                      type="number"
                      min="0.01"
                      step="0.01"
                      value={row.rate}
                      onChange={(event) => updateLineItemRow(row.id, "rate", event.target.value)}
                    />
                  </label>
                  <label className="line-item-qty">
                    <span>Freight Amount</span>
                    <input
                      type="number"
                      min="0"
                      step="0.01"
                      value={row.transport}
                      onChange={(event) => updateLineItemRow(row.id, "transport", event.target.value)}
                    />
                  </label>
                  <label className="line-item-qty">
                    <span>GST Amount</span>
                    <input
                      type="number"
                      min="0"
                      step="0.01"
                      value={row.gstPercent}
                      onChange={(event) => updateLineItemRow(row.id, "gstPercent", event.target.value)}
                    />
                  </label>
                  <label className="line-item-total">
                    <span>Total amount</span>
                    <input
                      type="number"
                      readOnly
                      step="0.01"
                      value={row.productId ? formatAmountInput(amounts.totalAmount) : ""}
                      placeholder={product ? formatAmountInput(amounts.computedTotalAmount) : "Select product first"}
                    />
                  </label>
                  <button
                    className="icon-action-button danger"
                    type="button"
                    onClick={() => removeLineItemRow(row.id)}
                    title="Delete line item"
                  >
                    🗑
                  </button>
                  <div className="line-item-preview">
                    <span>{product?.narration || product?.name || "Product details will auto-fill from the catalog."}</span>
                    <strong>{product ? formatCurrency(amounts.totalAmount) : "Amount pending"}</strong>
                    {product ? (
                      <>
                        <span>{`Base ${formatCurrency(amounts.unitValue)} | Freight ${formatCurrency(amounts.freightAmount)}`}</span>
                        <span>{`GST Amount ${formatCurrency(amounts.gstAmount)}`}</span>
                      </>
                    ) : null}
                  </div>
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

    const totalEnquiries = operations.totals?.enquiries || operations.enquiries.length;
    const dashboardMetrics = [
      { label: "Enquiries", value: totalEnquiries },
      ...operations.metrics.filter((metric) => metric.label !== "New Enquiries")
    ];
    const growthPoints: ChartPoint[] = [
      { label: "Enq", value: totalEnquiries },
      {
        label: "Draft",
        value: operations.metrics.find((item) => item.label === "Draft Quote")?.value ?? 0
      },
      {
        label: "Parsed",
        value: operations.enquiries.filter((item) => item.parserStatus === "Parsed").length
      },
      {
        label: "Approved",
        value: operations.metrics.find((item) => item.label === "Approved Quote")?.value ?? 0
      },
      { label: "Sent", value: operations.metrics.find((item) => item.label === "Sent Quote")?.value ?? 0 },
      { label: "Orders", value: operations.metrics.find((item) => item.label === "Orders")?.value ?? 0 }
    ];

    const totalQuotations = operations.totals?.quotations || operations.quotations.length;
    const sentQuotations = operations.quotations.filter((item) => item.status === "Sent Quote").length;

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
              {operations.actions.defaultTemplateFolderUrl ? (
                <a href={operations.actions.defaultTemplateFolderUrl} target="_blank" rel="noreferrer">
                  Default Templates
                </a>
              ) : null}
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
          {dashboardMetrics.map((metric) => (
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
                label: "Total enquiries",
                value: String(totalEnquiries)
              },
              {
                label: "Parsed",
                value: String(operations.enquiries.filter((item) => item.parserStatus === "Parsed").length)
              },
              {
                label: "Draft quote",
                value: String(
                  operations.enquiries.filter((item) => item.parserStatus === "Draft Quote").length
                )
              }
            ]}
          />
          <DetailCard
            title="Quotation Pulse"
            rows={[
              { label: "Draft records", value: String(operations.totals?.quotations || operations.quotations.length) },
              {
                label: "Approved quote",
                value: String(
                  operations.quotations.filter((item) => item.status === "Approved Quote").length
                )
              },
              {
                label: "Sent",
                value: String(
                  operations.quotations.filter((item) => item.status === "Sent Quote").length
                )
              }
            ]}
          />
          <DetailCard
            title="Business Pulse"
            rows={[
              { label: "Customers", value: String(operations.totals?.customers || operations.customers.length) },
              { label: "Products", value: String(operations.totals?.products || operations.products.length) },
              { label: "Orders", value: String(operations.totals?.orders || operations.orders.length) }
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

    const enquiryActionState =
      actionState &&
      (actionState.key.startsWith("enquiry-draft-") ||
        actionState.key.startsWith("line-items-") ||
        actionState.key === "portal-enquiry" ||
        actionState.key === "product-documents" ||
        actionState.key === "send-product-documents")
        ? actionState
        : null;

    return (
      <>
        {enquiryActionState ? (
          <section className={`action-banner ${enquiryActionState.status}`}>
            <strong>{enquiryActionState.label}</strong>
            <span>{enquiryActionState.message}</span>
            {renderBannerCloseButton(dismissActionState)}
          </section>
        ) : null}
        {enquiryPage.error ? (
          <section className="action-banner error">
            <strong>Enquiries</strong>
            <span>{enquiryPage.error}</span>
            {renderBannerCloseButton(() => setEnquiryPage((current) => ({ ...current, error: "" })))}
          </section>
        ) : null}
      <PaginatedTable
        eyebrow="Enquiries"
        title="All inbound requests from Airtable forms"
        headerAction={
          <div className="table-header-actions">
            <label className="table-filter-inline">
              <span>Status</span>
              <select
                value={enquiryStatusFilter}
                onChange={(event) => {
                  const nextStatus = event.target.value;
                  setEnquiryStatusFilter(nextStatus);
                  setEnquiryPage({
                    ...createPagedState<EnquiryRecord>(),
                    loading: true
                  });
                  void loadEnquiryPage("", "reset", nextStatus);
                }}
              >
                {enquiryStatusOptions.map((statusOption) => (
                  <option key={statusOption} value={statusOption}>
                    {statusOption}
                  </option>
                ))}
              </select>
            </label>
            <button className="action-inline-button" type="button" onClick={openEnquiryEntry}>
              Create Enquiry
            </button>
          </div>
        }
        rows={filteredEnquiries}
        columns={
          <>
            <th>Enquiry</th>
            <th>Lead</th>
            <th>Status</th>
            <th>Customer</th>
            <th>Quotation</th>
            <th>Requirement</th>
            <th>Folder</th>
            <th>Action</th>
          </>
        }
        emptyTitle={enquiryPage.loading ? "Loading enquiries..." : "No enquiries yet"}
        emptyBody={enquiryPage.loading ? "Fetching the current enquiry page." : "New Airtable form submissions will appear here."}
        pagination={{
          canNext: Boolean(enquiryPage.nextOffset),
          canPrevious: enquiryPage.previousOffsets.length > 0,
          isLoading: enquiryPage.loading,
          label: `Page ${enquiryPage.page}`,
          metaLabel: `${filteredEnquiries.length} on this page${enquiryPage.totalCount ? ` of ${enquiryPage.totalCount}` : ""}`,
          onNext: () => {
            if (enquiryPage.nextOffset) {
              void loadEnquiryPage(enquiryPage.nextOffset, "next");
            }
          },
          onPrevious: () =>
            void loadEnquiryPage(enquiryPage.previousOffsets[enquiryPage.previousOffsets.length - 1] || "", "previous")
        }}
        renderRow={(enquiry) => {
          const customer = customerLookup.get(enquiry.linkedCustomerId);
          const quotation = enquiry.quotations.length
            ? quotationLookup.get(enquiry.quotations[0])
            : undefined;

          return (
            <tr key={enquiry.id}>
              <td>
                <strong>{enquiry.enquiryId}</strong>
                {enquiry.loggedDateTime ? (
                  <span className="table-submeta">Logged {formatDateTime(enquiry.loggedDateTime)}</span>
                ) : null}
              </td>
              <td>
                <strong>{enquiry.leadName || "Unnamed lead"}</strong>
                <span>{enquiry.phone || enquiry.email || "Contact missing"}</span>
                <span className="table-submeta">{enquiry.state || "State pending"}</span>
              </td>
              <td>
                <span className={`status-chip ${statusTone(enquiry.parserStatus)}`}>
                  {enquiry.parserStatus || "New"}
                </span>
              </td>
              <td>{customer?.customerName || "Not linked yet"}</td>
              <td>{quotation?.quotationNumber || "Not created"}</td>
              <td>
                <strong>{enquiry.requirementSummary || "Product not selected"}</strong>
              </td>
              <td>
                {enquiry.driveFolderUrl ? (
                  <a
                    className="action-inline-link"
                    href={enquiry.driveFolderUrl}
                    target="_blank"
                    rel="noreferrer"
                    onClick={() => handleLinkAction(`folder-${enquiry.id}`, "Open Folder")}
                    title="Open Drive folder"
                  >
                    ↗
                  </a>
                ) : (
                  "Not created"
                )}
              </td>
              <td>
                <div className="action-stack">
                  <button
                    className="icon-action-button neutral"
                    type="button"
                    onClick={() => openEnquiryEdit(enquiry)}
                    title="Edit enquiry"
                  >
                    ✎
                  </button>
                  {["New", "New Enquiries", "Parsed"].includes(enquiry.parserStatus || "") ? (
                    <button
                      className="action-inline-button"
                      type="button"
                      onClick={() => void handleGenerateDraftFromEnquiry(enquiry.id)}
                      disabled={actionState?.key === `enquiry-draft-${enquiry.id}` && actionState.status === "loading"}
                    >
                      {actionState?.key === `enquiry-draft-${enquiry.id}` && actionState.status === "loading"
                        ? "Generating..."
                        : "Generate Draft"}
                    </button>
                  ) : null}
                  {(!enquiry.linkedCustomerId || !enquiry.quotations.length) &&
                  (enquiry.parserStatus === "New" || enquiry.parserStatus === "Parsed") ? (
                    <button
                      className="action-inline-button"
                      type="button"
                      onClick={() => void handleCreateCustomer(enquiry.id)}
                      disabled={actionState?.key === `create-customer-${enquiry.id}` && actionState.status === "loading"}
                    >
                      {actionState?.key === `create-customer-${enquiry.id}` && actionState.status === "loading"
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
                  ) : enquiry.parserStatus === "Draft Quote" && enquiry.driveFolderUrl ? (
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
      </>
    );
  };

  const renderCustomers = () => {
    if (!operations) {
      return null;
    }

    const customerRows = customerPage.initialized ? customerPage.rows : operations.customers;
    const previousOffset = customerPage.previousOffsets[customerPage.previousOffsets.length - 1] || "";

    return (
      <>
      {customerPage.error ? (
        <section className="action-banner error">
          <strong>Customers</strong>
          <span>{customerPage.error}</span>
          {renderBannerCloseButton(() => setCustomerPage((current) => ({ ...current, error: "" })))}
        </section>
      ) : null}
      <PaginatedTable
        eyebrow="Customers"
        title="Master account records generated from enquiry intake"
        headerAction={
          <button className="action-inline-button" type="button" onClick={() => openCustomerEntry()}>
            New Customer
          </button>
        }
        rows={customerRows}
        columns={
          <>
            <th>Client ID</th>
            <th>Customer</th>
            <th>Company</th>
            <th>Contact</th>
            <th>Type</th>
            <th>Drive</th>
            <th>Action</th>
          </>
        }
        emptyTitle={customerPage.loading ? "Loading customers..." : "No customers yet"}
        emptyBody={
          customerPage.loading
            ? "Fetching the current customer page."
            : "Customer records created from enquiries will appear here."
        }
        pagination={{
          canNext: Boolean(customerPage.nextOffset),
          canPrevious: customerPage.previousOffsets.length > 0,
          isLoading: customerPage.loading,
          label: `Page ${customerPage.page}`,
          metaLabel: `${customerRows.length} on this page${customerPage.totalCount ? ` of ${customerPage.totalCount}` : ""}`,
          onNext: () => {
            if (customerPage.nextOffset) {
              void loadCustomerPage(customerPage.nextOffset, "next");
            }
          },
          onPrevious: () => void loadCustomerPage(previousOffset, "previous")
        }}
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
            <td>
              <button className="icon-action-button neutral" type="button" title="Edit customer" onClick={() => openCustomerEntry(customer)}>
                ✎
              </button>
            </td>
          </tr>
        )}
      />
      </>
    );
  };

  const renderQuotationActionButtons = (quotation: QuotationRecord) => {
    const whatsappActionKey = `quotation-send-whatsapp-${quotation.id}`;
    const regenerateActionKey = `quotation-regenerate-${quotation.id}`;
    const markSentActionKey = `quotation-mark-sent-${quotation.id}`;
    const isWhatsAppLoading =
      actionState?.key === whatsappActionKey && actionState.status === "loading";
    const isRegenerateLoading =
      actionState?.key === regenerateActionKey && actionState.status === "loading";
    const isMarkingSent =
      actionState?.key === markSentActionKey && actionState.status === "loading";
    const existingOrder = orderByQuotationId.get(quotation.id);

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

    if (quotation.status === "Approved Quote") {
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
              className="icon-action-button whatsapp"
              type="button"
              title="Send quotation on WhatsApp"
              onClick={() => void handleSendQuotation(quotation.id, "whatsapp")}
              disabled={isWhatsAppLoading}
            >
              {isWhatsAppLoading ? "..." : "WA"}
            </button>
            <button
              className="icon-action-button neutral"
              type="button"
              title="Mark quotation as sent"
              onClick={() => void handleMarkQuotationSent(quotation.id)}
              disabled={isMarkingSent}
            >
              {isMarkingSent ? "..." : "S"}
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

    if (quotation.status === "Sent Quote") {
      return (
        <div className="action-stack">
          <div className="action-icon-row">
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
            <button
              className="icon-action-button neutral"
              type="button"
              title={existingOrder ? "Edit linked order" : "Create order"}
              onClick={() => openOrderEntry(existingOrder, quotation)}
            >
              {existingOrder ? "O" : "+"}
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

    if (!quotation.draftFileUrl && !quotation.finalPdfUrl) {
      return (
        <div className="action-stack">
          <button
            className="action-inline-button"
            type="button"
            onClick={() => void handleRegenerateQuotationDraft(quotation.id)}
            disabled={isRegenerateLoading}
          >
            {isRegenerateLoading ? "Generating..." : "Generate Draft"}
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

    if (
      quotation.draftFileUrl ||
      quotation.status === "Draft Quote" ||
      quotation.status === "Parsed" ||
      quotation.status === "New Enquiries"
    ) {
      return (
        <div className="action-stack">
          <div className="action-icon-row">
            <button
              className="icon-action-button neutral"
              type="button"
              title="Add or edit line items"
              onClick={() => {
                openLineItemEntry(quotation.id);
                setActionState({
                  key: `quotation-line-items-${quotation.id}`,
                  label: `Edit Line Items for ${quotation.quotationNumber}`,
                  status: "success",
                  message: "Line item entry is ready in the portal."
                });
              }}
            >
              LI
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
            <button
              className="icon-action-button neutral"
              type="button"
              title="Generate final PDF"
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

    return <span>{formatDate(quotation.sentDate)}</span>;
  };

  const renderQuotations = (title: string, subtitle: string, statuses: string[]) => {
    if (!operations) {
      return null;
    }

    const currentQuotationKey = statuses.join("|") || "all";
    const filteredQuotations =
      quotationPage.initialized && quotationPageKey === currentQuotationKey
        ? quotationPage.rows
        : operations.quotations.filter((quotation) =>
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
        {quotationPage.error && quotationPageKey === currentQuotationKey ? (
          <section className="action-banner error">
            <strong>Quotations</strong>
            <span>{quotationPage.error}</span>
            {renderBannerCloseButton(() => setQuotationPage((current) => ({ ...current, error: "" })))}
          </section>
        ) : null}
        <PaginatedTable
          eyebrow="Quotations"
          title={title}
          subtitle={subtitle}
          headerAction={
              statuses.includes("Draft Quote") || statuses.includes("Parsed") || statuses.includes("New Enquiries")
              ? (
                <button className="action-inline-button" type="button" onClick={openQuotationEntry}>
                  Create Quotation
                </button>
              )
              : undefined
          }
          rows={filteredQuotations}
          columns={
            <>
              <th>Quotation</th>
              <th>Customer</th>
              <th>Status</th>
            <th>Items</th>
            {statuses.length === 1 && statuses[0] === "Approved Quote" ? null : <th>Draft</th>}
            <th>Final PDF</th>
            <th>Actions</th>
            </>
          }
          emptyTitle="No quotations in this section"
          emptyBody="As the workflow progresses, matching quotation records will appear here."
          pagination={{
            canNext: Boolean(quotationPage.nextOffset),
            canPrevious: quotationPage.previousOffsets.length > 0,
            isLoading: quotationPage.loading,
            label: `Page ${quotationPage.page}`,
            metaLabel: `${filteredQuotations.length} on this page${quotationPage.totalCount ? ` of ${quotationPage.totalCount}` : ""}`,
            onNext: () => {
              if (quotationPage.nextOffset) {
                void loadQuotationPage(quotationPage.nextOffset, "next", statuses);
              }
            },
            onPrevious: () =>
              void loadQuotationPage(
                quotationPage.previousOffsets[quotationPage.previousOffsets.length - 1] || "",
                "previous",
                statuses
              )
          }}
          renderRow={(quotation) => (
            <tr key={quotation.id}>
              <td>
                <strong>{quotation.quotationNumber}</strong>
                <span>
                  {enquiryLookup.get(quotation.linkedEnquiryId)?.enquiryId ||
                    quotation.referenceNumber ||
                    quotation.linkedEnquiryId ||
                    "No enquiry link"}
                </span>
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
              <td>{quotation.lineItemCount || 0}</td>
              {statuses.length === 1 && statuses[0] === "Approved Quote" ? null : (
                <td>
                  {quotation.status === "Approved Quote" ? (
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
            <th>Status</th>
            <th>Order Date</th>
            <th>Value</th>
            <th>Payment</th>
            <th>Action</th>
          </>
        }
        emptyTitle="No orders yet"
        emptyBody="Accepted quotations will surface here as orders."
        renderRow={(order) => (
          <tr key={order.id}>
            <td>{order.orderNumber}</td>
            <td>{customerLookup.get(order.linkedCustomerId)?.customerName || "Not linked"}</td>
            <td>{quotationLookup.get(order.linkedQuotationId)?.quotationNumber || "Not linked"}</td>
            <td>{order.orderStatus || "Pending"}</td>
            <td>{formatDate(order.orderDate)}</td>
            <td>{(order.totalAmount || order.orderValue) ? formatCurrency(order.totalAmount || order.orderValue) : "Pending"}</td>
            <td>{order.paymentStatus || "Pending"}</td>
            <td>
              <button className="icon-action-button neutral" type="button" title="Edit order" onClick={() => openOrderEntry(order)}>
                ✎
              </button>
            </td>
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
          <ReshamSutraLogo className="auth-logo" />
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
          <ReshamSutraLogo className="auth-logo" />
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
          ["Sent Quote"]
        );
        break;
      case "approvedQuotations":
        content = renderQuotations(
          "Approved quotations ready for dispatch",
          "Quotes that have cleared internal review and are waiting for send or final conversion.",
          ["Approved Quote"]
        );
        break;
      case "quotationDrafts":
        content = renderQuotations(
          "Quotation drafts prepared from live enquiries",
          "Internal working drafts that still need line item and commercial review.",
          ["Draft Quote", "Parsed", "New Enquiries"]
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
            <ReshamSutraLogo className="brand-logo" />
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
