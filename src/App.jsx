import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Factory,
  CalendarDays,
  FileSpreadsheet,
  Download,
  LayoutDashboard,
  Flame,
  Truck,
  Warehouse,
  AlertTriangle,
  Container,
  ClipboardCheck,
  Copy,
  Upload,
  Plus,
  Search,
  X,
  FileUp,
  UploadCloud,
  PackagePlus,
  ArrowUpRight,
} from 'lucide-react';

// ============================================================================
// 1. TIỆN ÍCH NGÀY THÁNG & DANH SÁCH TUẦN NĂM 2026 (THỨ 2 -> THỨ 7)
// ============================================================================
const pad2 = (n) => (n < 10 ? `0${n}` : `${n}`);

const formatDM = (d) => `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}`;

const formatDMY = (d) =>
  `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;

const fmtNum = (n) => {
  if (n === null || n === undefined || Number.isNaN(Number(n))) return '0';
  return Number(n).toLocaleString('en-US', { maximumFractionDigits: 1 });
};

function generateWeeks2026() {
  const weeks = [];
  const baseMonday = new Date(2025, 11, 29); // Thứ 2 đầu Tuần 1/2026
  const dayNames = ['THỨ 2', 'THỨ 3', 'THỨ 4', 'THỨ 5', 'THỨ 6', 'THỨ 7'];

  for (let w = 1; w <= 53; w++) {
    const mon = new Date(baseMonday);
    mon.setDate(baseMonday.getDate() + (w - 1) * 7);
    const sat = new Date(mon);
    sat.setDate(mon.getDate() + 5);

    if (mon.getFullYear() > 2026) break;

    const days = dayNames.map((dayCode, idx) => {
      const cur = new Date(mon);
      cur.setDate(mon.getDate() + idx);
      return {
        dayCode,
        dateDM: formatDM(cur),
        fullLabel: `${dayCode} (${formatDM(cur)})`,
      };
    });

    weeks.push({
      weekNum: w,
      monStr: formatDMY(mon),
      satStr: formatDMY(sat),
      label: `Tuần ${pad2(w)} (Thứ 2 ${formatDMY(mon)} - Thứ 7 ${formatDMY(sat)})`,
      days,
    });
  }
  return weeks;
}

const WEEKS_2026 = generateWeeks2026();

// ============================================================================
// 2. DỮ LIỆU MẪU (TUẦN 48 & TỒN KHO TUẦN 46)
// ============================================================================
const BASE_PRODUCTION_48 = [
  { stt: 1, dept: 'Lò đốt chất thải', shift: 'Ca sáng', staff: '—', days: 'Không hoạt động', output: 'Ngưng hoạt động (Lý do: không có nguyên liệu)', supervisor: '—', active: false },
  { stt: 1, dept: 'Lò đốt chất thải', shift: 'Ca đêm', staff: 'Tổ vận hành lò', days: '26, 27, 28/11', output: 'Tổng: 10,000 kg/ca (30,000 kg/tuần)', supervisor: 'Huỳnh Tấn Tuấn', active: true },
  { stt: 2, dept: 'Hệ thống tái chế dung môi', shift: 'Ca sáng & đêm', staff: '—', days: 'Không hoạt động', output: 'Ngưng hoạt động (Lý do: không có nguyên liệu)', supervisor: '—', active: false },
  { stt: 3, dept: 'Hệ thống nước thải', shift: 'Ca sáng', staff: 'Giang', days: '23, 24, 25, 26, 27, 28/11', output: 'Tổng: 10 m³/ngày (60 m³/tuần)', supervisor: 'Huỳnh Tấn Tuấn', active: true },
  { stt: 4, dept: 'Hệ thống máy đóng gạch', shift: 'Ca sáng & đêm', staff: '—', days: 'Không hoạt động', output: 'Ngưng hoạt động (Lý do: không có nguyên liệu)', supervisor: '—', active: false },
  { stt: 5, dept: 'Hệ thống xúc rửa thùng phi', shift: 'Ca sáng & đêm', staff: '—', days: 'Không hoạt động', output: 'Ngưng hoạt động (Lý do: không có nguyên liệu)', supervisor: '—', active: false },
  { stt: 6, dept: 'Phân loại, lấy hàng', shift: 'Ca sáng', staff: 'Vân, Niên, Cường, Hiếu', days: '23, 24, 25, 26, 27, 28/11', output: 'Phân loại rác về trong ngày 23–28/11', supervisor: 'Huỳnh Tấn Tuấn', active: true },
  { stt: 7, dept: 'Chuyển giao chất thải', shift: 'Ca sáng', staff: 'Tuấn', days: '23, 25, 27, 28/11', output: 'Chuyển giao cho KBEC, Đất sạch, ATY ECO', supervisor: 'Huỳnh Tấn Tuấn', active: true },
  { stt: 8, dept: 'Giám sát Nội bộ', shift: 'Ca sáng', staff: 'Hồ Tân An', days: '23, 24, 25, 26, 27, 28/11', output: 'Giám sát toàn bộ hoạt động của nhà máy', supervisor: 'Hồ Tân An', active: true },
];

const BASE_APPENDIX1_48 = [
  { source: 'CÔNG TY TNHH ARAKAWA CHEMICAL VIỆT NAM', waste: 'Bao bì mềm thải', code: '18 01 01', date: '26/11', qty: 6000 },
  { source: 'CÔNG TY TNHH ARAKAWA CHEMICAL VIỆT NAM', waste: 'Gỗ thải nhiễm các TPNH', code: '12 08 01', date: '26/11', qty: 300 },
  { source: 'CTNH Nội bộ', waste: 'Bùn từ HTXLKT', code: '12 01 03', date: '23–28/11', qty: 147 },
  { source: 'CTNH Nội bộ', waste: 'Bùn từ HTXL nước thải', code: '12 02 02', date: '23–28/11', qty: 3413 },
  { source: 'CTCN Nội bộ', waste: 'Rác Công nghiệp nội bộ', code: '-', date: 'Tháng 11', qty: 200 },
  { source: 'CTSH Nội bộ', waste: 'Rác thải sinh hoạt nội bộ', code: '-', date: 'Tháng 11', qty: 325 },
];

const BASE_APPENDIX2_48 = [
  { type: 'Chất thải công nghiệp thông thường', qty: '100,000 kg (100 tấn)', receiver: 'KBEC, Đất sạch, ATY ECO', schedule: '23, 25, 27, 28/11', pic: 'Huỳnh Tấn Tuấn' },
  { type: 'Bùn thải công nghiệp thông thường', qty: '60,000 kg (60 tấn)', receiver: 'KBEC (Bùn Gamuda Land)', schedule: '23, 25, 27/11', pic: 'Nguyễn Trung Thanh Hòa' },
];

const BASE_COLLECTION_48 = [
  { stt: 1, dayCode: 'THỨ 2', company: 'Công ty Cổ phần Sản xuất Kinh doanh XNK Vĩnh Long', address: 'Lô A1, Khu C, KCN Hoà Phú, Xã Phú Quới, Vĩnh Long', waste: 'CTNH', weight: 12000, carrier: 'ĐỖ CAO TRÍ', note: 'XIN 2 NHÂN CÔNG HUY THỊNH', truckStatus: 'CÓ LỊCH CHẠY' },
  { stt: 2, dayCode: 'THỨ 2', company: 'CTY TNHH LON NƯỚC GIẢI KHÁT TBC BALL', address: 'Số 2 VSIP 2-A đường số 18, KCN VN-Singapore, Vĩnh Tân', waste: 'BÙN + RÁC', weight: 15000, carrier: 'HUY THỊNH', note: 'XIN 1 NHÂN CÔNG HUY THỊNH', truckStatus: 'CÓ LỊCH CHẠY' },
  { stt: 3, dayCode: 'THỨ 2', company: 'CÔNG TY TNHH TÂN MINH TRÍ', address: 'Phân hiệu trường TH Quang Trung - 276 Lê Lợi, P. Tam Thắng, HCM', waste: 'XÀ BẦN', weight: 10000, carrier: 'ĐỖ CAO TRÍ', note: '', truckStatus: '—' },
  { stt: 4, dayCode: 'THỨ 2', company: 'CÔNG TY TNHH XÂY DỰNG TAM THẮNG', address: 'Trường TH Quang Trung - 28 đường 30/4, P. Tam Thắng, TP.HCM', waste: 'XÀ BẦN', weight: 0, carrier: 'CHƯA CHỐT', note: 'Chưa cập nhật KL', truckStatus: '—' },
  { stt: 5, dayCode: 'THỨ 2', company: 'CTCP TẬP ĐOÀN HÀ ĐÔ - CN KHÁCH SẠN IBIS HÀ ĐÔ', address: '2 Hồng Hà, Phường Tân Sơn Hoà, TP. Hồ Chí Minh', waste: 'CTNH', weight: 2000, carrier: 'NINH GIA', note: '', truckStatus: '—' },
  { stt: '-', dayCode: 'THỨ 3', company: 'CÔNG TY NGHỈ HOẠT ĐỘNG THU GOM', address: 'Toàn tuyến nghỉ theo kế hoạch', waste: 'NGHỈ', weight: 0, carrier: 'NGHỈ', note: 'Nghỉ Thứ 3', truckStatus: 'NGHỈ' },
  { stt: 6, dayCode: 'THỨ 4', company: 'CÔNG TY TNHH TIGER TECH VIỆT NAM', address: 'Lô G5A, đường số 3, KCN Việt Hương 2, phường Tây Nam', waste: 'CTNH & CTCN', weight: 10000, carrier: 'ĐỖ CAO TRÍ', note: '', truckStatus: 'CHƯA CÓ LỊCH CHẠY' },
  { stt: 7, dayCode: 'THỨ 4', company: 'CÔNG TY CP GIẢI PHÁP THANH TOÁN VIỆT NAM', address: 'Toà Central Park - 117 Nguyễn Du, P. Bến Thành, Q.1, HCM', waste: 'CTNH', weight: 1000, carrier: 'NINH GIA', note: '', truckStatus: '—' },
  { stt: 8, dayCode: 'THỨ 5', company: 'CÔNG TY CỔ PHẦN LIWAYWAY VIỆT NAM', address: 'Số 14 đường số 5, KCN Việt Nam-Singapore, phường Bình Hoà', waste: 'XÀ BẦN', weight: 25000, carrier: 'HUY THỊNH', note: '', truckStatus: 'CÓ LỊCH CHẠY' },
  { stt: 9, dayCode: 'THỨ 5', company: 'CÔNG TY TNHH ARAKAWA CHEMICAL VIỆT NAM', address: 'Lô D1 Đường D3, KCN chuyên sâu Phú Mỹ 3, P. Tân Phước, TP.HCM', waste: 'CTNH + CTCN', weight: 8000, carrier: 'ĐỖ CAO TRÍ', note: 'XIN 2 NHÂN CÔNG HUY THỊNH (CẦN 6 TỚI 8 TANK CẮT NẮP)', truckStatus: '—' },
  { stt: 10, dayCode: 'THỨ 6', company: 'CTY TNHH LON NƯỚC GIẢI KHÁT TBC BALL', address: 'Số 2 VSIP 2-A đường số 18, KCN VN-Singapore, Vĩnh Tân', waste: 'BÙN + RÁC', weight: 16000, carrier: 'HUY THỊNH', note: 'TANK CẮT NẮP CỦA TBC', truckStatus: 'DỰ KIẾN CÓ LỊCH CHẠY' },
  { stt: 11, dayCode: 'THỨ 6', company: 'CÔNG TY CỔ PHẦN MICROM', address: 'Đường số 10, KCN Cơ Khí Ô Tô, Xã Bình Mỹ, TP.HCM', waste: 'BÙN NGUY HẠI + CTNH', weight: 6000, carrier: 'ĐỖ CAO TRÍ', note: 'CHUẨN BỊ 10 TANK CẮT NẮP', truckStatus: '—' },
  { stt: '-', dayCode: 'THỨ 7', company: 'CHƯA CÓ LỊCH CHẠY', address: 'Chờ điều phối bổ sung', waste: '—', weight: 0, carrier: 'CHƯA CHỐT', note: '', truckStatus: 'CHƯA CÓ LỊCH CHẠY' },
];

const BASE_INVENTORY_ALL = [
  { warehouse: 'CTNH', date: '06/11/2026', customer: 'CTCP PHÂN BÓN VÀ HÓA CHẤT DẦU KHÍ ĐÔNG NAM BỘ', code: 'CTNH', inQty: 20, outQty: 0, note: '' },
  { warehouse: 'CTNH', date: '07/11/2026', customer: 'CÔNG TY TNHH THƯƠNG MẠI VÀ SẢN XUẤT DT', code: 'CTNH', inQty: 6410, outQty: 0, note: '' },
  { warehouse: 'CTNH', date: '09/11/2026', customer: 'CÔNG TY TNHH ZHIHE', code: 'CTNH', inQty: 3560, outQty: 0, note: '' },
  { warehouse: 'CTNH', date: '09/11/2026', customer: 'CÔNG TY TNHH VINMOS (Lô 1)', code: 'CTNH', inQty: 0, outQty: 9490, note: 'Chưa xử lý' },
  { warehouse: 'CTNH', date: '09/11/2026', customer: 'CÔNG TY TNHH VINMOS (Lô 2)', code: 'CTNH', inQty: 0, outQty: 29480, note: 'Chưa xử lý' },
  { warehouse: 'CTNH', date: '11/11/2026', customer: 'CÔNG TY TNHH LỐP ADVANCE VIỆT NAM', code: 'CTNH', inQty: 9650, outQty: 0, note: 'Đã bán nhớt và xử lý cặn' },
  { warehouse: 'CTNH', date: '12/11/2026', customer: 'CÔNG TY TNHH THƯƠNG MẠI KHÔI MINH', code: 'CTNH', inQty: 5767, outQty: 0, note: '' },
  { warehouse: 'Bùn', date: '06/11/2026', customer: 'CÔNG TY TNHH ĐẤT SẠCH VIỆT', code: 'BTCN', inQty: 0, outQty: 38530, note: 'Xuất chuyển giao ĐẤT SẠCH' },
  { warehouse: 'Bùn', date: '10/11/2026', customer: 'CÔNG TY CỔ PHẦN PHÁT TRIỂN HẠ TẦNG VIỆT SƠN', code: 'BTNH', inQty: 28760, outQty: 0, note: '' },
  { warehouse: 'Bùn', date: '12/11/2026', customer: 'CÔNG TY CỔ PHẦN GIẤY HOÀNG HÀ PHÚ YÊN', code: 'BTCN', inQty: 52210, outQty: 0, note: 'Bùn đã đem chôn KBEC & chuyển giao' },
  { warehouse: 'Tiêu hủy', date: '15/09/2026', customer: 'CÔNG TY TNHH DƯỢC PHẨM SNN VIỆT NAM', code: 'Hàng hóa tiêu hủy', inQty: 17170, outQty: 0, note: 'Chờ phân loại' },
  { warehouse: 'Tiêu hủy', date: '16/09/2026', customer: 'CÔNG TY TNHH AGRONALIS', code: 'Hàng hóa tiêu hủy', inQty: 22000, outQty: 0, note: 'Bột dùng thấm nước (Tồn 5,000)' },
  { warehouse: 'Tiêu hủy', date: '01/10/2026', customer: 'CÔNG TY CỔ PHẦN CHARME PERFUME', code: 'Chất thải thông thường', inQty: 22200, outQty: 0, note: 'Chờ phân loại' },
  { warehouse: 'Tiêu hủy', date: '14/11/2026', customer: 'CÔNG TY TNHH FISHY VIỆT NAM', code: 'TIEUHUY (CTCN)', inQty: 12474, outQty: 0, note: 'Đã hủy' },
  { warehouse: 'CTCNTT', date: '09/11/2026', customer: 'CÔNG TY TNHH THIẾT KẾ AKI', code: 'CTCN', inQty: 1250, outQty: 0, note: 'Đã chuyển giao KBEC' },
  { warehouse: 'CTCNTT', date: '13/11/2026', customer: 'CÔNG TY TNHH LON NƯỚC GIẢI KHÁT TBC-BALL VIỆT NAM', code: 'CTCN', inQty: 350, outQty: 0, note: 'Đã chuyển giao KBEC' },
  { warehouse: 'Nước thải', date: '18/09/2026', customer: 'CÔNG TY TNHH DEGO HOLDING', code: 'NTCN', inQty: 48960, outQty: 0, note: '' },
  { warehouse: 'Nước thải', date: '11/11/2026', customer: 'CÔNG TY CỔ PHẦN NĂNG LƯỢNG HT', code: 'NTNH', inQty: 53790, outQty: 0, note: 'Đã xử lý 18 tank (Xả 41 tank 22–24/10)' },
  { warehouse: 'Nước thải', date: '12/11/2026', customer: 'CÔNG TY TNHH ASIA PACKAGING INDUSTRIES VIỆT NAM', code: 'NTNH', inQty: 25300, outQty: 0, note: '' },
];

const emptyWeekPayload = () => ({
  production: [],
  appendix1: [],
  appendix2: [],
  collection: [],
  inventory: [],
  transferKg: 0,
  incineratorKg: 0,
});

function buildDefaultStore() {
  return {
    48: {
      production: structuredClone(BASE_PRODUCTION_48),
      appendix1: structuredClone(BASE_APPENDIX1_48),
      appendix2: structuredClone(BASE_APPENDIX2_48),
      collection: structuredClone(BASE_COLLECTION_48),
      inventory: structuredClone(BASE_INVENTORY_ALL),
      transferKg: 160000,
      incineratorKg: 30000,
    },
    46: {
      ...emptyWeekPayload(),
      inventory: structuredClone(BASE_INVENTORY_ALL),
    },
  };
}

// ============================================================================
// 3. COMPONENT CHÍNH CỦA ỨNG DỤNG REACT
// ============================================================================
export default function App() {
  const [currentWeekNum, setCurrentWeekNum] = useState(48);
  const [activeTab, setActiveTab] = useState('overview');

  const [weeklyStore, setWeeklyStore] = useState(() => {
    try {
      const saved = localStorage.getItem('huythinh_react_store_v1');
      return saved ? JSON.parse(saved) : buildDefaultStore();
    } catch {
      return buildDefaultStore();
    }
  });

  // Bộ lọc Tab Thu gom & Tồn kho
  const [selectedDay, setSelectedDay] = useState('ALL');
  const [selectedCarrier, setSelectedCarrier] = useState('ALL');
  const [selectedWarehouse, setSelectedWarehouse] = useState('ALL');
  const [inventorySearch, setInventorySearch] = useState('');
  const [inventoryStatus, setInventoryStatus] = useState('ALL');

  // Modal Nhập Excel & Thêm chuyến
  const [isImportOpen, setIsImportOpen] = useState(false);
  const [importType, setImportType] = useState('collection');
  const [selectedFile, setSelectedFile] = useState(null);
  const fileInputRef = useRef(null);

  const [isAddTripOpen, setIsAddTripOpen] = useState(false);
  const [newTrip, setNewTrip] = useState({
    dayCode: 'THỨ 2',
    company: '',
    address: '',
    waste: 'CTNH',
    weight: '',
    carrier: 'HUY THỊNH',
    note: '',
  });

  useEffect(() => {
    try {
      localStorage.setItem('huythinh_react_store_v1', JSON.stringify(weeklyStore));
    } catch {
      // Bỏ qua nếu vượt giới hạn storage
    }
  }, [weeklyStore]);

  const weekMeta = useMemo(
    () => WEEKS_2026.find((w) => w.weekNum === Number(currentWeekNum)) || WEEKS_2026[47],
    [currentWeekNum]
  );

  const weekData = useMemo(
    () => weeklyStore[currentWeekNum] || emptyWeekPayload(),
    [weeklyStore, currentWeekNum]
  );

  const wLabel = `Tuần ${pad2(weekMeta.weekNum)}`;
  const hasAnyData =
    weekData.collection.length > 0 ||
    weekData.inventory.length > 0 ||
    weekData.production.length > 0;

  // Các chỉ số tổng hợp cho Tab Tổng quan
  const totalCollectionKg = useMemo(
    () => weekData.collection.reduce((sum, r) => sum + (Number(r.weight) || 0), 0),
    [weekData]
  );

  const totalInventoryInKg = useMemo(
    () => weekData.inventory.reduce((sum, r) => sum + (Number(r.inQty) || 0), 0),
    [weekData]
  );

  const dailyChartRows = useMemo(() => {
    return weekMeta.days.map((d) => {
      const huyThinh = weekData.collection
        .filter((r) => r.dayCode === d.dayCode && r.carrier === 'HUY THỊNH')
        .reduce((s, r) => s + (Number(r.weight) || 0), 0);
      const doCaoTri = weekData.collection
        .filter((r) => r.dayCode === d.dayCode && r.carrier === 'ĐỖ CAO TRÍ')
        .reduce((s, r) => s + (Number(r.weight) || 0), 0);
      const ninhGia = weekData.collection
        .filter((r) => r.dayCode === d.dayCode && r.carrier === 'NINH GIA')
        .reduce((s, r) => s + (Number(r.weight) || 0), 0);
      const total = huyThinh + doCaoTri + ninhGia;
      return { ...d, huyThinh, doCaoTri, ninhGia, total };
    });
  }, [weekMeta, weekData]);

  const maxDayTotal = useMemo(
    () => Math.max(1, ...dailyChartRows.map((r) => r.total)),
    [dailyChartRows]
  );

  const warehouseChartRows = useMemo(() => {
    const names = ['CTNH', 'Bùn', 'Tiêu hủy', 'CTCNTT', 'Nước thải'];
    const rows = names.map((name) => {
      const inKg = weekData.inventory
        .filter((r) => r.warehouse === name)
        .reduce((s, r) => s + (Number(r.inQty) || 0), 0);
      const outKg = weekData.inventory
        .filter((r) => r.warehouse === name)
        .reduce((s, r) => s + (Number(r.outQty) || 0), 0);
      return { name, inTons: +(inKg / 1000).toFixed(2), outTons: +(outKg / 1000).toFixed(2) };
    });
    const maxTons = Math.max(1, ...rows.flatMap((r) => [r.inTons, r.outTons]));
    return { rows, maxTons };
  }, [weekData]);

  // Danh sách thu gom đã lọc
  const filteredCollection = useMemo(() => {
    return weekData.collection.filter((r) => {
      const matchDay = selectedDay === 'ALL' || r.dayCode === selectedDay;
      const matchCarrier = selectedCarrier === 'ALL' || r.carrier === selectedCarrier;
      return matchDay && matchCarrier;
    });
  }, [weekData, selectedDay, selectedCarrier]);

  // Danh sách tồn kho đã lọc
  const filteredInventory = useMemo(() => {
    const q = inventorySearch.trim().toLowerCase();
    return weekData.inventory.filter((r) => {
      const matchWh = selectedWarehouse === 'ALL' || r.warehouse === selectedWarehouse;
      const matchQ =
        !q ||
        r.customer?.toLowerCase().includes(q) ||
        r.code?.toLowerCase().includes(q) ||
        r.note?.toLowerCase().includes(q);
      const noteLower = (r.note || '').toLowerCase();
      const isPending = noteLower.includes('chờ') || noteLower.includes('chưa');
      const isProcessed =
        noteLower.includes('đã hủy') ||
        noteLower.includes('đã chuyển') ||
        noteLower.includes('đã xử lý') ||
        noteLower.includes('đã bán') ||
        Number(r.outQty) > 0;

      if (inventoryStatus === 'PENDING') return matchWh && matchQ && isPending;
      if (inventoryStatus === 'PROCESSED') return matchWh && matchQ && isProcessed;
      return matchWh && matchQ;
    });
  }, [weekData, selectedWarehouse, inventorySearch, inventoryStatus]);

  // ==========================================================================
  // CÁC HÀNH ĐỘNG CẬP NHẬT DỮ LIỆU
  // ==========================================================================
  const handleChangeWeek = (val) => {
    setCurrentWeekNum(Number(val));
    setSelectedDay('ALL');
    setSelectedWarehouse('ALL');
  };

  const handleCloneWeek48 = () => {
    const source = weeklyStore[48];
    if (!source) return;
    setWeeklyStore((prev) => ({
      ...prev,
      [currentWeekNum]: structuredClone(source),
    }));
  };

  const handleSaveNewTrip = (e) => {
    e.preventDefault();
    if (!newTrip.company.trim()) return;

    setWeeklyStore((prev) => {
      const target = prev[currentWeekNum] ? structuredClone(prev[currentWeekNum]) : emptyWeekPayload();
      const nextStt = target.collection.filter((r) => typeof r.stt === 'number').length + 1;
      target.collection.push({
        stt: nextStt,
        dayCode: newTrip.dayCode,
        company: newTrip.company.trim(),
        address: newTrip.address.trim() || 'Chưa cập nhật địa chỉ',
        waste: newTrip.waste.trim() || 'CTNH',
        weight: Number(newTrip.weight) || 0,
        carrier: newTrip.carrier,
        note: newTrip.note.trim(),
        truckStatus: 'CÓ LỊCH CHẠY',
      });
      return { ...prev, [currentWeekNum]: target };
    });

    setNewTrip({
      dayCode: weekMeta.days[0]?.dayCode || 'THỨ 2',
      company: '',
      address: '',
      waste: 'CTNH',
      weight: '',
      carrier: 'HUY THỊNH',
      note: '',
    });
    setIsAddTripOpen(false);
  };

  const handleProcessExcel = () => {
    if (!selectedFile) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buf = new Uint8Array(evt.target.result);
        const wb = XLSX.read(buf, { type: 'array' });

        if (importType === 'collection') {
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
          let activeDay = 'THỨ 2';
          const parsed = [];

          rows.forEach((r) => {
            if (!r || r.length < 3) return;
            const dayText = String(r[1] || '').toUpperCase();
            if (dayText.includes('THỨ 2')) activeDay = 'THỨ 2';
            else if (dayText.includes('THỨ 3')) activeDay = 'THỨ 3';
            else if (dayText.includes('THỨ 4')) activeDay = 'THỨ 4';
            else if (dayText.includes('THỨ 5')) activeDay = 'THỨ 5';
            else if (dayText.includes('THỨ 6')) activeDay = 'THỨ 6';
            else if (dayText.includes('THỨ 7')) activeDay = 'THỨ 7';

            const company = String(r[2] || '').trim();
            if (!company || company === 'CÔNG TY') return;

            parsed.push({
              stt: r[0] || '-',
              dayCode: activeDay,
              company,
              address: String(r[3] || ''),
              waste: String(r[4] || ''),
              weight: parseFloat(String(r[5] || '0').replace(/,/g, '')) || 0,
              carrier: String(r[6] || 'CHƯA CHỐT').toUpperCase(),
              note: String(r[7] || ''),
              truckStatus: String(r[8] || '—'),
            });
          });

          setWeeklyStore((prev) => {
            const target = prev[currentWeekNum] ? structuredClone(prev[currentWeekNum]) : emptyWeekPayload();
            target.collection = parsed;
            return { ...prev, [currentWeekNum]: target };
          });
        } else {
          const parsedInv = [];
          wb.SheetNames.forEach((sName) => {
            const sheet = wb.Sheets[sName];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            rows.forEach((r, idx) => {
              if (idx < 9 || !r) return;
              const col0 = String(r[0] || '');
              if (col0.includes('Tổng') || col0.includes('Người lập') || col0.includes('MISA')) return;
              const customer = String(r[2] || r[4] || '').trim();
              const inQty = parseFloat(r[6] || r[4] || 0) || 0;
              const outQty = parseFloat(r[9] || r[7] || 0) || 0;
              if (customer && (inQty > 0 || outQty > 0)) {
                parsedInv.push({
                  warehouse: sName,
                  date: String(r[0] || r[1] || '').slice(0, 10),
                  customer,
                  code: String(r[1] || r[3] || sName),
                  inQty,
                  outQty,
                  note: String(r[7] || r[8] || ''),
                });
              }
            });
          });

          setWeeklyStore((prev) => {
            const target = prev[currentWeekNum] ? structuredClone(prev[currentWeekNum]) : emptyWeekPayload();
            target.inventory = parsedInv;
            return { ...prev, [currentWeekNum]: target };
          });
        }

        setSelectedFile(null);
        setIsImportOpen(false);
      } catch (err) {
        alert(`Không thể đọc tệp Excel: ${err.message}`);
      }
    };
    reader.readAsArrayBuffer(selectedFile);
  };

  const handleExportCsv = () => {
    let csv = `\uFEFFBAO CAO VAN HANH HUY THINH - ${weekMeta.label}\n`;
    csv += 'Ngay,Khach hang,Loai chat thai,Khoi luong (kg),Nha xe,Ghi chu\n';
    weekData.collection.forEach((r) => {
      csv += `"${r.dayCode}","${r.company}","${r.waste}",${r.weight},"${r.carrier}","${r.note}"\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Huy_Thinh_Tuan_${pad2(weekMeta.weekNum)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleResetDefault = () => {
    localStorage.removeItem('huythinh_react_store_v1');
    setWeeklyStore(buildDefaultStore());
    setCurrentWeekNum(48);
    setSelectedFile(null);
    setIsImportOpen(false);
  };

  // ==========================================================================
  // GIAO DIỆN
  // ==========================================================================
  return (
    <div className="min-h-screen flex flex-col bg-slate-100 text-slate-900">
      {/* HEADER + COMBOBOX TUẦN TOÀN CỤC */}
      <header className="bg-slate-900 text-white border-b border-slate-800 sticky top-0 z-30 shadow-md">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-emerald-500/20 border border-emerald-400/30 flex items-center justify-center text-emerald-400 shrink-0">
              <Factory className="w-6 h-6" />
            </div>
            <div>
              <div className="flex flex-wrap items-center gap-2">
                <h1 className="text-base sm:text-lg font-bold tracking-tight">
                  CÔNG TY TNHH THƯƠNG MẠI HUY THỊNH
                </h1>
                <span className="px-2 py-0.5 text-[11px] font-semibold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 rounded-full">
                  Tóc Tiên • Châu Pha
                </span>
              </div>
              <p className="text-xs text-slate-400">
                Hệ thống Điều hành Sản xuất, Kế hoạch Thu gom & Thống kê Tồn kho
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <div className="flex items-center gap-2 bg-slate-800 border border-slate-700 rounded-lg px-2.5 py-1.5">
              <CalendarDays className="w-4 h-4 text-emerald-400 shrink-0" />
              <label htmlFor="week-combo" className="text-xs font-medium text-slate-300 hidden sm:inline">
                Kỳ báo cáo:
              </label>
              <select
                id="week-combo"
                value={currentWeekNum}
                onChange={(e) => handleChangeWeek(e.target.value)}
                className="bg-transparent text-xs sm:text-sm font-semibold text-white focus:outline-none cursor-pointer"
              >
                {WEEKS_2026.map((w) => {
                  const d = weeklyStore[w.weekNum];
                  const flag =
                    d && (d.collection.length > 0 || d.inventory.length > 0 || d.production.length > 0)
                      ? ' [Có dữ liệu]'
                      : '';
                  return (
                    <option key={w.weekNum} value={w.weekNum} className="bg-slate-900 text-white">
                      {w.label}
                      {flag}
                    </option>
                  );
                })}
              </select>
            </div>

            <button
              onClick={() => setIsImportOpen(true)}
              className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white transition"
            >
              <FileSpreadsheet className="w-4 h-4" />
              Nhập Excel / Cập nhật
            </button>

            <button
              onClick={handleExportCsv}
              className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition"
            >
              <Download className="w-4 h-4" />
              Xuất CSV
            </button>
          </div>
        </div>

        {/* THANH 4 TAB */}
        <nav className="bg-white border-b border-slate-200 text-slate-700">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-2 flex overflow-x-auto gap-1.5">
            {[
              { id: 'overview', label: '1. Tổng Quan Điều Hành', Icon: LayoutDashboard },
              { id: 'production', label: `2. Lệnh Sản Xuất (${wLabel})`, Icon: Flame },
              { id: 'collection', label: `3. Kế Hoạch Thu Gom (${wLabel})`, Icon: Truck },
              { id: 'inventory', label: `4. Thống Kê Tồn Kho (${wLabel})`, Icon: Warehouse },
            ].map(({ id, label, Icon }) => {
              const isActive = activeTab === id;
              return (
                <button
                  key={id}
                  onClick={() => setActiveTab(id)}
                  className={`inline-flex items-center gap-2 px-3.5 py-2 rounded-lg text-xs sm:text-sm font-semibold whitespace-nowrap transition ${
                    isActive
                      ? 'bg-slate-900 text-white shadow-sm'
                      : 'text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  <Icon className="w-4 h-4" />
                  {label}
                </button>
              );
            })}
          </div>
        </nav>
      </header>

      {/* NỘI DUNG CHÍNH */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
        {/* BANNER TRẠNG THÁI TUẦN */}
        <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-lg bg-slate-900 text-white flex items-center justify-center font-bold text-sm">
              T{pad2(weekMeta.weekNum)}
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-900">Đang xem dữ liệu: {weekMeta.label}</h2>
              <p className="text-xs text-slate-500">
                {hasAnyData
                  ? `Dữ liệu vận hành từ Thứ 2 (${weekMeta.monStr}) đến Thứ 7 (${weekMeta.satStr}) đồng bộ trên cả 4 tab.`
                  : `Chưa có dữ liệu cho ${wLabel}. Bạn có thể sao chép mẫu từ Tuần 48 hoặc nhập file Excel.`}
              </p>
            </div>
          </div>

          {!hasAnyData && (
            <div className="flex items-center gap-2">
              <button
                onClick={handleCloneWeek48}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-amber-50 text-amber-800 border border-amber-300 hover:bg-amber-100 transition"
              >
                <Copy className="w-3.5 h-3.5" />
                Sao chép mẫu từ Tuần 48
              </button>
              <button
                onClick={() => setIsImportOpen(true)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-emerald-600 text-white hover:bg-emerald-500 transition"
              >
                <Upload className="w-3.5 h-3.5" />
                Nhập Excel cho tuần này
              </button>
            </div>
          )}
        </div>

        {/* TAB 1: TỔNG QUAN */}
        {activeTab === 'overview' && (
          <div className="space-y-6">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                    Thu Gom Dự Kiến ({wLabel})
                  </span>
                  <div className="w-9 h-9 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center">
                    <Truck className="w-5 h-5" />
                  </div>
                </div>
                <div className="mt-3 flex items-baseline gap-2">
                  <span className="text-2xl font-extrabold text-slate-900">{fmtNum(totalCollectionKg)}</span>
                  <span className="text-xs font-semibold text-slate-500">
                    kg ({(totalCollectionKg / 1000).toFixed(1)} tấn)
                  </span>
                </div>
                <p className="mt-2 text-xs text-slate-600">
                  {weekData.collection.filter((r) => r.weight > 0).length} điểm thu gom trong tuần
                </p>
              </div>

              <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                    Tổng Nhập Tồn Kho ({wLabel})
                  </span>
                  <div className="w-9 h-9 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center">
                    <PackagePlus className="w-5 h-5" />
                  </div>
                </div>
                <div className="mt-3 flex items-baseline gap-2">
                  <span className="text-2xl font-extrabold text-slate-900">{fmtNum(totalInventoryInKg)}</span>
                  <span className="text-xs font-semibold text-slate-500">
                    kg (~{(totalInventoryInKg / 1000).toFixed(1)} tấn)
                  </span>
                </div>
                <p className="mt-2 text-xs text-slate-600">5 phân hệ kho: CTNH, Bùn, Tiêu hủy, CTCN, Nước thải</p>
              </div>

              <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                    Chuyển Giao Ngoài ({wLabel})
                  </span>
                  <div className="w-9 h-9 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center">
                    <ArrowUpRight className="w-5 h-5" />
                  </div>
                </div>
                <div className="mt-3 flex items-baseline gap-2">
                  <span className="text-2xl font-extrabold text-slate-900">{fmtNum(weekData.transferKg)}</span>
                  <span className="text-xs font-semibold text-slate-500">
                    kg ({(weekData.transferKg / 1000).toFixed(0)} tấn)
                  </span>
                </div>
                <p className="mt-2 text-xs text-slate-600">KBEC, Đất sạch, ATY ECO</p>
              </div>

              <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                    Công Suất Lò Đốt
                  </span>
                  <div className="w-9 h-9 rounded-xl bg-rose-50 text-rose-600 flex items-center justify-center">
                    <Flame className="w-5 h-5" />
                  </div>
                </div>
                <div className="mt-3 flex items-baseline gap-2">
                  <span className="text-2xl font-extrabold text-slate-900">{fmtNum(weekData.incineratorKg)}</span>
                  <span className="text-xs font-semibold text-slate-500">kg</span>
                </div>
                <p className="mt-2 text-xs text-slate-600">10,000 kg/ca • HT Nước thải 10 m³/ngày</p>
              </div>
            </div>

            {/* BIỂU ĐỒ TRỰC QUAN THU GOM THEO NGÀY & CƠ CẤU KHO */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              <div className="lg:col-span-7 bg-white rounded-2xl p-5 border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h3 className="text-base font-bold text-slate-900">
                      Thu Gom Theo Ngày — {wLabel} (Thứ 2 – Thứ 7)
                    </h3>
                    <p className="text-xs text-slate-500">Đơn vị: tấn (Phân bổ theo nhà xe)</p>
                  </div>
                  <div className="flex items-center gap-3 text-[11px] font-semibold">
                    <span className="flex items-center gap-1 text-emerald-700">
                      <i className="w-2.5 h-2.5 rounded-sm bg-emerald-500 inline-block" /> Huy Thịnh
                    </span>
                    <span className="flex items-center gap-1 text-blue-700">
                      <i className="w-2.5 h-2.5 rounded-sm bg-blue-500 inline-block" /> Đỗ Cao Trí
                    </span>
                    <span className="flex items-center gap-1 text-amber-700">
                      <i className="w-2.5 h-2.5 rounded-sm bg-amber-500 inline-block" /> Ninh Gia
                    </span>
                  </div>
                </div>

                <div className="space-y-3">
                  {dailyChartRows.map((d) => {
                    const pct = Math.round((d.total / maxDayTotal) * 100);
                    return (
                      <div key={d.dayCode} className="space-y-1">
                        <div className="flex justify-between text-xs font-semibold">
                          <span className="text-slate-700">{d.fullLabel}</span>
                          <span className="text-slate-900">{(d.total / 1000).toFixed(1)} tấn</span>
                        </div>
                        <div className="w-full h-3 bg-slate-100 rounded-full overflow-hidden flex">
                          <div
                            style={{ width: `${(d.huyThinh / maxDayTotal) * 100}%` }}
                            className="bg-emerald-500 h-full"
                            title={`Huy Thịnh: ${fmtNum(d.huyThinh)} kg`}
                          />
                          <div
                            style={{ width: `${(d.doCaoTri / maxDayTotal) * 100}%` }}
                            className="bg-blue-500 h-full"
                            title={`Đỗ Cao Trí: ${fmtNum(d.doCaoTri)} kg`}
                          />
                          <div
                            style={{ width: `${(d.ninhGia / maxDayTotal) * 100}%` }}
                            className="bg-amber-500 h-full"
                            title={`Ninh Gia: ${fmtNum(d.ninhGia)} kg`}
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="lg:col-span-5 bg-white rounded-2xl p-5 border border-slate-200 shadow-sm">
                <h3 className="text-base font-bold text-slate-900 mb-1">Cơ Cấu Nhập & Xuất 5 Kho</h3>
                <p className="text-xs text-slate-500 mb-4">Đơn vị: tấn</p>
                <div className="space-y-3.5">
                  {warehouseChartRows.rows.map((w) => (
                    <div key={w.name} className="space-y-1">
                      <div className="flex justify-between text-xs font-semibold">
                        <span className="text-slate-800">{w.name}</span>
                        <span className="text-slate-600">
                          Nhập: <strong>{w.inTons}T</strong> • Xuất: <strong className="text-rose-600">{w.outTons}T</strong>
                        </span>
                      </div>
                      <div className="w-full h-2.5 bg-slate-100 rounded-full overflow-hidden">
                        <div
                          style={{ width: `${Math.min(100, Math.round((w.inTons / warehouseChartRows.maxTons) * 100))}%` }}
                          className="h-full bg-slate-900 rounded-full"
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* CẢNH BÁO VẬN HÀNH */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
              <div className="bg-amber-50/80 border border-amber-200 rounded-2xl p-4 flex gap-3.5">
                <div className="w-10 h-10 rounded-xl bg-amber-100 text-amber-700 flex items-center justify-center shrink-0">
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="text-sm font-bold text-amber-900">4 Hệ Thống Ngưng Do Thiếu Nguyên Liệu</h4>
                  <p className="text-xs text-amber-800 mt-1">
                    Lò đốt ca sáng, Tái chế dung môi, Máy đóng gạch và Xúc rửa thùng phi ngưng chạy khi chưa có nguyên liệu.
                  </p>
                </div>
              </div>

              <div className="bg-blue-50/80 border border-blue-200 rounded-2xl p-4 flex gap-3.5">
                <div className="w-10 h-10 rounded-xl bg-blue-100 text-blue-700 flex items-center justify-center shrink-0">
                  <Container className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="text-sm font-bold text-blue-900">Tank Cắt Nắp & Nhân Công Bốc Xếp</h4>
                  <p className="text-xs text-blue-800 mt-1">
                    Arakawa cần 6–8 tank + 2 nhân công; Microm cần 10 tank; XNK Vĩnh Long xin 2 nhân công; TBC xin 1 nhân công.
                  </p>
                </div>
              </div>

              <div className="bg-rose-50/80 border border-rose-200 rounded-2xl p-4 flex gap-3.5">
                <div className="w-10 h-10 rounded-xl bg-rose-100 text-rose-700 flex items-center justify-center shrink-0">
                  <ClipboardCheck className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="text-sm font-bold text-rose-900">Điểm Nóng Tồn Kho Cần Theo Dõi</h4>
                  <p className="text-xs text-rose-800 mt-1">
                    Hàng chờ phân loại tại kho Tiêu hủy, tồn Bùn cuối kỳ và các lô CTNH chưa xử lý cần rà soát tiến độ.
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: LỆNH SẢN XUẤT */}
        {activeTab === 'production' && (
          <div className="space-y-6">
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="p-5 border-b border-slate-200 bg-slate-50/60">
                <span className="text-xs font-bold uppercase tracking-wider text-emerald-600">
                  Lệnh Sản Xuất • {weekMeta.label}
                </span>
                <h2 className="text-lg font-bold text-slate-900 mt-0.5">
                  Kế Hoạch Vận Hành Hệ Thống & Phân Công Nhân Sự
                </h2>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-sm">
                  <thead>
                    <tr className="bg-slate-100 text-slate-600 text-xs uppercase border-b border-slate-200">
                      <th className="py-3 px-4 font-bold">STT</th>
                      <th className="py-3 px-4 font-bold">Bộ phận</th>
                      <th className="py-3 px-4 font-bold">Ca</th>
                      <th className="py-3 px-4 font-bold">Người thực hiện</th>
                      <th className="py-3 px-4 font-bold">Thời gian</th>
                      <th className="py-3 px-4 font-bold">Sản lượng / Trạng thái</th>
                      <th className="py-3 px-4 font-bold">Giám sát</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {weekData.production.length === 0 ? (
                      <tr>
                        <td colSpan="7" className="py-10 text-center text-slate-400">
                          Chưa có Lệnh sản xuất cho {wLabel}.
                        </td>
                      </tr>
                    ) : (
                      weekData.production.map((row, i) => (
                        <tr key={`${row.dept}-${i}`} className={row.active ? 'hover:bg-slate-50' : 'bg-slate-50/70 text-slate-500'}>
                          <td className="py-3 px-4 font-bold">{row.stt}</td>
                          <td className="py-3 px-4 font-bold text-slate-900">{row.dept}</td>
                          <td className="py-3 px-4 text-xs">{row.shift}</td>
                          <td className="py-3 px-4">{row.staff}</td>
                          <td className="py-3 px-4 text-xs font-semibold">{row.days}</td>
                          <td className="py-3 px-4">
                            <span
                              className={`inline-block px-2 py-0.5 rounded-full text-xs font-semibold mr-2 ${
                                row.active ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-600'
                              }`}
                            >
                              {row.active ? 'Hoạt động' : 'Ngưng'}
                            </span>
                            <span className="text-xs">{row.output}</span>
                          </td>
                          <td className="py-3 px-4 text-xs font-semibold">{row.supervisor}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              <div className="lg:col-span-7 bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="p-4 border-b border-slate-200 bg-rose-50/50 flex items-center justify-between">
                  <h3 className="text-base font-bold text-slate-900">Phụ Lục 01 — Chất Thải Đưa Vào Lò Đốt</h3>
                  <span className="text-xs font-bold px-2.5 py-1 rounded-lg bg-rose-100 text-rose-800">
                    Tổng: {fmtNum(weekData.appendix1.reduce((s, r) => s + (Number(r.qty) || 0), 0))} kg
                  </span>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-50 text-slate-600 uppercase border-b border-slate-200">
                        <th className="py-2.5 px-3 font-bold">Nguồn</th>
                        <th className="py-2.5 px-3 font-bold">Tên chất thải</th>
                        <th className="py-2.5 px-3 font-bold">Mã</th>
                        <th className="py-2.5 px-3 font-bold">Ngày</th>
                        <th className="py-2.5 px-3 font-bold text-right">KL (kg)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {weekData.appendix1.map((r, i) => (
                        <tr key={i}>
                          <td className="py-2.5 px-3 font-medium">{r.source}</td>
                          <td className="py-2.5 px-3">{r.waste}</td>
                          <td className="py-2.5 px-3 font-mono">{r.code}</td>
                          <td className="py-2.5 px-3">{r.date}</td>
                          <td className="py-2.5 px-3 text-right font-bold">{fmtNum(r.qty)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="lg:col-span-5 bg-white rounded-2xl border border-slate-200 shadow-sm p-4 space-y-3">
                <h3 className="text-base font-bold text-slate-900">Phụ Lục 02 — Chuyển Giao Ngoài Nhà Máy</h3>
                {weekData.appendix2.map((r, i) => (
                  <div key={i} className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-1">
                    <div className="flex justify-between items-center">
                      <span className="text-xs font-bold text-emerald-700">{r.type}</span>
                      <span className="text-xs font-extrabold text-slate-900">{r.qty}</span>
                    </div>
                    <p className="text-xs text-slate-600">Đơn vị nhận: {r.receiver}</p>
                    <p className="text-xs text-slate-500">
                      Ngày: {r.schedule} • Phụ trách: {r.pic}
                    </p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* TAB 3: KẾ HOẠCH THU GOM */}
        {activeTab === 'collection' && (
          <div className="space-y-6">
            <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  onClick={() => setSelectedDay('ALL')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold ${
                    selectedDay === 'ALL' ? 'bg-slate-900 text-white' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                  }`}
                >
                  Tất cả (Thứ 2 - Thứ 7)
                </button>
                {weekMeta.days.map((d) => (
                  <button
                    key={d.dayCode}
                    onClick={() => setSelectedDay(d.dayCode)}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold ${
                      selectedDay === d.dayCode
                        ? 'bg-slate-900 text-white'
                        : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                    }`}
                  >
                    {d.fullLabel}
                  </button>
                ))}
              </div>

              <div className="flex items-center gap-2">
                <select
                  value={selectedCarrier}
                  onChange={(e) => setSelectedCarrier(e.target.value)}
                  className="text-xs font-semibold border border-slate-300 rounded-lg px-3 py-2 bg-white"
                >
                  <option value="ALL">Tất cả Nhà xe</option>
                  <option value="HUY THỊNH">Xe Huy Thịnh</option>
                  <option value="ĐỖ CAO TRÍ">Nhà xe Đỗ Cao Trí</option>
                  <option value="NINH GIA">Nhà xe Ninh Gia</option>
                </select>

                <button
                  onClick={() => setIsAddTripOpen(true)}
                  className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold bg-blue-600 hover:bg-blue-500 text-white"
                >
                  <Plus className="w-4 h-4" />
                  Thêm chuyến
                </button>
              </div>
            </div>

            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-sm">
                  <thead>
                    <tr className="bg-slate-100 text-slate-600 text-xs uppercase border-b border-slate-200">
                      <th className="py-3 px-4 font-bold">STT</th>
                      <th className="py-3 px-4 font-bold">Ngày</th>
                      <th className="py-3 px-4 font-bold">Khách hàng / Địa chỉ</th>
                      <th className="py-3 px-4 font-bold">Chất thải</th>
                      <th className="py-3 px-4 font-bold text-right">KL (kg)</th>
                      <th className="py-3 px-4 font-bold">Nhà xe</th>
                      <th className="py-3 px-4 font-bold">Ghi chú</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {filteredCollection.length === 0 ? (
                      <tr>
                        <td colSpan="7" className="py-10 text-center text-slate-400">
                          Không có chuyến thu gom nào phù hợp.
                        </td>
                      </tr>
                    ) : (
                      filteredCollection.map((r, idx) => {
                        const dObj = weekMeta.days.find((d) => d.dayCode === r.dayCode);
                        return (
                          <tr key={idx} className="hover:bg-slate-50">
                            <td className="py-3 px-4 font-bold">{r.stt}</td>
                            <td className="py-3 px-4 text-xs font-bold whitespace-nowrap">
                              {dObj ? dObj.fullLabel : r.dayCode}
                            </td>
                            <td className="py-3 px-4">
                              <div className="font-bold text-slate-900 text-xs sm:text-sm">{r.company}</div>
                              <div className="text-xs text-slate-500">{r.address}</div>
                            </td>
                            <td className="py-3 px-4 text-xs font-semibold">{r.waste}</td>
                            <td className="py-3 px-4 text-right font-extrabold">
                              {r.weight > 0 ? fmtNum(r.weight) : '—'}
                            </td>
                            <td className="py-3 px-4">
                              <span className="px-2.5 py-1 rounded-lg text-xs font-bold bg-slate-100 text-slate-800">
                                {r.carrier}
                              </span>
                            </td>
                            <td className="py-3 px-4 text-xs text-slate-600">{r.note || '—'}</td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TAB 4: THỐNG KÊ TỒN KHO */}
        {activeTab === 'inventory' && (
          <div className="space-y-6">
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
              {[
                { key: 'ALL', label: 'Tất cả 5 Kho' },
                { key: 'CTNH', label: 'Kho CTNH' },
                { key: 'Bùn', label: 'Kho Bùn' },
                { key: 'Tiêu hủy', label: 'Kho Tiêu Hủy' },
                { key: 'CTCNTT', label: 'Kho CTCNTT' },
                { key: 'Nước thải', label: 'Kho Nước Thải' },
              ].map((g) => {
                const rows =
                  g.key === 'ALL'
                    ? weekData.inventory
                    : weekData.inventory.filter((r) => r.warehouse === g.key);
                const sumIn = rows.reduce((s, r) => s + (Number(r.inQty) || 0), 0);
                const active = selectedWarehouse === g.key;
                return (
                  <button
                    key={g.key}
                    onClick={() => setSelectedWarehouse(g.key)}
                    className={`text-left p-3.5 rounded-2xl border transition ${
                      active
                        ? 'bg-slate-900 text-white border-slate-900 shadow-md'
                        : 'bg-white text-slate-800 border-slate-200 hover:border-slate-300'
                    }`}
                  >
                    <div className={`text-xs font-bold uppercase ${active ? 'text-emerald-400' : 'text-slate-500'}`}>
                      {g.label}
                    </div>
                    <div className="text-base font-extrabold mt-1">
                      {fmtNum(sumIn)} <span className="text-[11px] font-normal opacity-75">kg</span>
                    </div>
                    <div className={`text-[11px] mt-1 ${active ? 'text-slate-300' : 'text-slate-400'}`}>
                      {rows.length} dòng
                    </div>
                  </button>
                );
              })}
            </div>

            <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="relative flex-1 max-w-md">
                <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={inventorySearch}
                  onChange={(e) => setInventorySearch(e.target.value)}
                  placeholder="Tìm khách hàng, mã chất thải, ghi chú..."
                  className="w-full pl-10 pr-4 py-2 text-sm border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-emerald-500"
                />
              </div>

              <select
                value={inventoryStatus}
                onChange={(e) => setInventoryStatus(e.target.value)}
                className="text-xs font-semibold border border-slate-300 rounded-xl px-3 py-2 bg-white"
              >
                <option value="ALL">Tất cả trạng thái</option>
                <option value="PENDING">Chờ phân loại / Chưa xử lý</option>
                <option value="PROCESSED">Đã hủy / Đã chuyển giao / Đã xử lý</option>
              </select>
            </div>

            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto max-h-[520px]">
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="sticky top-0 bg-slate-100 text-slate-600 text-xs uppercase border-b border-slate-200">
                    <tr>
                      <th className="py-3 px-4 font-bold">Kho</th>
                      <th className="py-3 px-4 font-bold">Ngày</th>
                      <th className="py-3 px-4 font-bold">Khách hàng</th>
                      <th className="py-3 px-4 font-bold">Mã / Loại</th>
                      <th className="py-3 px-4 font-bold text-right">Nhập (kg)</th>
                      <th className="py-3 px-4 font-bold text-right">Xuất (kg)</th>
                      <th className="py-3 px-4 font-bold">Trạng thái / Ghi chú</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {filteredInventory.length === 0 ? (
                      <tr>
                        <td colSpan="7" className="py-10 text-center text-slate-400">
                          Không có bản ghi tồn kho phù hợp.
                        </td>
                      </tr>
                    ) : (
                      filteredInventory.map((r, idx) => (
                        <tr key={idx} className="hover:bg-slate-50">
                          <td className="py-3 px-4 text-xs font-bold">{r.warehouse}</td>
                          <td className="py-3 px-4 text-xs text-slate-600 whitespace-nowrap">{r.date}</td>
                          <td className="py-3 px-4 font-semibold text-xs sm:text-sm">{r.customer}</td>
                          <td className="py-3 px-4 text-xs font-mono text-slate-600">{r.code}</td>
                          <td className="py-3 px-4 text-right font-bold">
                            {r.inQty > 0 ? fmtNum(r.inQty) : '—'}
                          </td>
                          <td className="py-3 px-4 text-right font-bold text-rose-600">
                            {r.outQty > 0 ? fmtNum(r.outQty) : '—'}
                          </td>
                          <td className="py-3 px-4 text-xs text-slate-700">{r.note || '—'}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* MODAL NHẬP EXCEL */}
      {isImportOpen && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-lg w-full p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center">
                  <FileUp className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-900">Nhập Dữ Liệu Từ File Excel</h3>
                  <p className="text-xs text-slate-500">Áp dụng cho: {weekMeta.label}</p>
                </div>
              </div>
              <button onClick={() => setIsImportOpen(false)} className="text-slate-400 hover:text-slate-600">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3">
              <label className="block text-xs font-bold uppercase text-slate-600">Loại dữ liệu</label>
              <select
                value={importType}
                onChange={(e) => setImportType(e.target.value)}
                className="w-full text-sm border border-slate-300 rounded-xl px-3.5 py-2.5 bg-slate-50 font-medium"
              >
                <option value="collection">Kế hoạch thu gom tuần (.xlsx / .xls / .csv)</option>
                <option value="inventory">Tổng hợp tồn kho nhiều sheet (.xls / .xlsx)</option>
              </select>

              <div
                onClick={() => fileInputRef.current?.click()}
                className="border-2 border-dashed border-slate-300 hover:border-emerald-500 rounded-xl p-6 text-center cursor-pointer bg-slate-50/50"
              >
                <UploadCloud className="w-8 h-8 text-emerald-600 mx-auto mb-2" />
                <p className="text-sm font-semibold text-slate-700">Nhấp để chọn tệp Excel</p>
                <p className="text-xs text-slate-400 mt-1">
                  {selectedFile ? selectedFile.name : 'Chưa chọn tệp nào'}
                </p>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={(e) => setSelectedFile(e.target.files?.[0] || null)}
                />
              </div>
            </div>

            <div className="flex items-center justify-between pt-3 border-t border-slate-100">
              <button
                onClick={handleResetDefault}
                className="text-xs font-semibold text-rose-600 hover:underline"
              >
                Khôi phục dữ liệu gốc
              </button>
              <div className="flex gap-2">
                <button
                  onClick={() => setIsImportOpen(false)}
                  className="px-4 py-2 text-xs font-semibold rounded-xl bg-slate-100 text-slate-700"
                >
                  Hủy
                </button>
                <button
                  onClick={handleProcessExcel}
                  className="px-4 py-2 text-xs font-semibold rounded-xl bg-emerald-600 text-white hover:bg-emerald-500"
                >
                  Cập nhật
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL THÊM CHUYẾN THU GOM */}
      {isAddTripOpen && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4">
          <form
            onSubmit={handleSaveNewTrip}
            className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-slate-200 space-y-4"
          >
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-base font-bold text-slate-900">Thêm Chuyến Thu Gom — {wLabel}</h3>
              <button type="button" onClick={() => setIsAddTripOpen(false)} className="text-slate-400 hover:text-slate-600">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3 text-sm">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">Ngày (Thứ 2 – Thứ 7)</label>
                <select
                  value={newTrip.dayCode}
                  onChange={(e) => setNewTrip((p) => ({ ...p, dayCode: e.target.value }))}
                  className="w-full border border-slate-300 rounded-xl px-3 py-2"
                >
                  {weekMeta.days.map((d) => (
                    <option key={d.dayCode} value={d.dayCode}>
                      {d.fullLabel}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">Tên Công ty / Khách hàng</label>
                <input
                  type="text"
                  required
                  value={newTrip.company}
                  onChange={(e) => setNewTrip((p) => ({ ...p, company: e.target.value }))}
                  className="w-full border border-slate-300 rounded-xl px-3 py-2"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">Địa chỉ</label>
                <input
                  type="text"
                  value={newTrip.address}
                  onChange={(e) => setNewTrip((p) => ({ ...p, address: e.target.value }))}
                  className="w-full border border-slate-300 rounded-xl px-3 py-2"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-600 mb-1">Loại chất thải</label>
                  <input
                    type="text"
                    value={newTrip.waste}
                    onChange={(e) => setNewTrip((p) => ({ ...p, waste: e.target.value }))}
                    className="w-full border border-slate-300 rounded-xl px-3 py-2"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-600 mb-1">Khối lượng (kg)</label>
                  <input
                    type="number"
                    value={newTrip.weight}
                    onChange={(e) => setNewTrip((p) => ({ ...p, weight: e.target.value }))}
                    className="w-full border border-slate-300 rounded-xl px-3 py-2"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-600 mb-1">Nhà xe</label>
                  <select
                    value={newTrip.carrier}
                    onChange={(e) => setNewTrip((p) => ({ ...p, carrier: e.target.value }))}
                    className="w-full border border-slate-300 rounded-xl px-3 py-2"
                  >
                    <option value="HUY THỊNH">HUY THỊNH</option>
                    <option value="ĐỖ CAO TRÍ">ĐỖ CAO TRÍ</option>
                    <option value="NINH GIA">NINH GIA</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-600 mb-1">Ghi chú</label>
                  <input
                    type="text"
                    value={newTrip.note}
                    onChange={(e) => setNewTrip((p) => ({ ...p, note: e.target.value }))}
                    className="w-full border border-slate-300 rounded-xl px-3 py-2"
                  />
                </div>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-3 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setIsAddTripOpen(false)}
                className="px-4 py-2 text-xs font-semibold rounded-xl bg-slate-100 text-slate-700"
              >
                Hủy
              </button>
              <button
                type="submit"
                className="px-4 py-2 text-xs font-semibold rounded-xl bg-blue-600 text-white hover:bg-blue-500"
              >
                Lưu vào lịch
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
}