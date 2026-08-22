import React, { useState, useMemo, useEffect } from 'react';
import Fuse from 'fuse.js';
import * as XLSX from 'xlsx';
import { 
  LayoutGrid, 
  ClipboardList, 
  BarChart3, 
  Settings, 
  AlertTriangle, 
  Boxes,
  Droplets,
  WifiOff,
  Leaf, 
  Coffee, 
  Archive,
  ArrowRight,
  ArrowLeft,
  Save,
  CheckCircle2,
  ShieldCheck,
  PlusCircle,
  Plus,
  Info,
  Trash2,
  ChevronDown,
  ChevronRight,
  PlusCircle as PlusCircleIcon,
  Search,
  X,
  Check,
  Edit3,
  History,
  FileText,
  LogOut,
  User as UserIcon,
  Lock,
  DollarSign,
  ShoppingCart,
  AlertCircle,
  ArrowDownCircle,
  ArrowUpCircle,
  Minus,
  Bell,
  Cloud,
  TrendingUp,
  AlertOctagon,
  Clock
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer } from 'recharts';
import { db, auth } from './firebase';
import PullToRefresh from './components/PullToRefresh';
import SwipeableLogItem from './components/SwipeableLogItem';
import { usePushNotifications } from './hooks/usePushNotifications';
import { 
  collection, 
  doc, 
  onSnapshot, 
  setDoc, 
  updateDoc, 
  deleteDoc,
  query,
  where,
  writeBatch,
  getDocFromServer,
  getDocs,
  waitForPendingWrites,
  enableNetwork,
  increment
} from 'firebase/firestore';
import { signInAnonymously, onAuthStateChanged } from 'firebase/auth';
import { InventoryItem, User, Role } from './types';

type LogEntry = {
  id: string;
  docId?: string;
  name: string;
  type: string;
  amount?: number;
  oldStock?: number;
  newStock?: number;
  reason?: string;
  timestamp: string;
  user: string;
  category?: string;
  theoreticalStock?: number;
  actualStock?: number;
  unit?: string;
  importPrice?: number;
  oldValue?: number;
  newValue?: number;
  details?: string;
};

type Screen = 'overview' | 'audit' | 'details' | 'edit' | 'report' | 'inventory' | 'settings' | 'add' | 'login' | 'inbound' | 'outbound' | 'financial' | 'suppliers';
type AuditType = 'daily' | 'weekly' | 'monthly';

const IOSSpinner = () => (
  <div className="ios-spinner">
    {[...Array(12)].map((_, i) => (
      <div key={i} className="ios-spinner-bar" />
    ))}
  </div>
);

const formatDate = (date: Date) => {
  const d = date.getDate().toString().padStart(2, '0');
  const m = (date.getMonth() + 1).toString().padStart(2, '0');
  const y = date.getFullYear();
  const h = date.getHours().toString().padStart(2, '0');
  const min = date.getMinutes().toString().padStart(2, '0');
  const s = date.getSeconds().toString().padStart(2, '0');
  return `${d}/${m}/${y} - ${h}:${min}:${s}`;
};

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | null | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | null | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const user = auth.currentUser;
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: user?.uid,
      email: user?.email,
      emailVerified: user?.emailVerified,
      isAnonymous: user?.isAnonymous,
      tenantId: user?.tenantId,
      providerInfo: user?.providerData.map(p => ({
        providerId: p.providerId,
        displayName: p.displayName,
        email: p.email,
        photoUrl: p.photoURL
      })) || []
    },
    operationType,
    path
  }
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// Sub-components
const ConfirmModal = ({ isOpen, title, message, onConfirm, onClose }: any) => (
  <AnimatePresence>
    {isOpen && (
      <div className="fixed inset-0 z-[100] flex items-center justify-center p-6 bg-black/40 backdrop-blur-sm">
        <motion.div 
          initial={{ opacity: 0, scale: 0.9 }}
          animate={{ opacity: 1, scale: 1 }}
          exit={{ opacity: 0, scale: 0.9 }}
          className="bg-white rounded-[32px] p-8 w-full max-w-sm shadow-2xl"
        >
          <div className="w-16 h-16 bg-red-50 text-red-500 rounded-2xl flex items-center justify-center mb-6">
            <AlertCircle size={32} />
          </div>
          <h3 className="text-2xl font-bold text-[#1C1C1E] mb-2">{title}</h3>
          <p className="text-apple-gray text-sm font-medium mb-8 leading-relaxed">{message}</p>
          <div className="flex gap-4">
            <button 
              onClick={onClose}
              className="flex-1 py-4 rounded-2xl bg-black/5 text-[#1C1C1E] font-bold text-sm active:scale-95 transition-all"
            >
              Hủy bỏ
            </button>
            <button 
              onClick={() => {
                onConfirm();
                onClose();
              }}
              className="flex-1 py-4 rounded-2xl bg-red-500 text-white font-bold text-sm active:scale-95 transition-all shadow-lg shadow-red-500/20"
            >
              Xác nhận
            </button>
          </div>
        </motion.div>
      </div>
    )}
  </AnimatePresence>
);

const Sidebar = ({ isMenuOpen, setIsMenuOpen, currentUser, currentScreen, navigateTo, handleLogout }: any) => (
  <AnimatePresence>
    {isMenuOpen && (
      <>
        <motion.div 
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          onClick={() => setIsMenuOpen(false)}
          className="fixed inset-0 z-[80] bg-black/20 backdrop-blur-sm"
        />
        <motion.div 
          initial={{ x: '-100%' }}
          animate={{ x: 0 }}
          exit={{ x: '-100%' }}
          transition={{ type: 'spring', damping: 25, stiffness: 200 }}
          className="fixed top-0 left-0 bottom-0 w-72 z-[90] bg-white shadow-2xl flex flex-col p-8"
        >
          <div className="flex items-center gap-3 mb-12">
            <div className="w-12 h-12 bg-apple-blue rounded-2xl flex items-center justify-center shadow-lg shadow-apple-blue/20">
              <Leaf className="text-white w-6 h-6" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-[#1C1C1E]">Quản lí kho hàng</h1>
              <p className="text-[10px] font-bold text-apple-gray uppercase tracking-widest">Hệ thống quản lý</p>
            </div>
          </div>

          <nav className="flex-1 space-y-2">
            {[
              { id: 'overview', label: 'Tổng quan', icon: LayoutGrid, managerOnly: true },
              { id: 'inventory', label: 'Danh mục kho', icon: Boxes },
              { id: 'inbound', label: 'Nhập kho', icon: ArrowDownCircle },
              { id: 'outbound', label: 'Xuất kho', icon: ArrowUpCircle },
              { id: 'audit', label: 'Kiểm kê', icon: ClipboardList },
              { id: 'report', label: 'Báo cáo', icon: BarChart3, managerOnly: true },
              { id: 'financial', label: 'Tài chính', icon: DollarSign, managerOnly: true },
              { id: 'suppliers', label: 'Mua hàng', icon: ShoppingCart, managerOnly: true },
              { id: 'settings', label: 'Cài đặt', icon: Settings },
            ].map((item) => {
              if (item.managerOnly && currentUser?.role !== 'quản lí') return null;
              const Icon = item.icon;
              return (
                <button
                  key={item.id}
                  onClick={() => {
                    navigateTo(item.id as Screen);
                    setIsMenuOpen(false);
                  }}
                  className={`w-full flex items-center gap-4 p-4 rounded-2xl transition-all ${
                    currentScreen === item.id 
                      ? 'bg-apple-blue text-white shadow-lg shadow-apple-blue/20' 
                      : 'text-apple-gray hover:bg-black/5'
                  }`}
                >
                  <Icon size={20} />
                  <span className="font-bold text-sm">{item.label}</span>
                </button>
              );
            })}
          </nav>

          <div className="pt-8 border-t border-black/5">
            <div className="flex items-center gap-3 mb-6">
              <div className="w-10 h-10 rounded-full bg-black/5 flex items-center justify-center overflow-hidden">
                <UserIcon size={20} className="text-apple-gray" />
              </div>
              <div>
                <p className="text-sm font-bold text-[#1C1C1E]">{currentUser?.name}</p>
                <p className="text-[10px] font-bold text-apple-gray uppercase tracking-widest">{currentUser?.role}</p>
              </div>
            </div>
            <button 
              onClick={() => {
                handleLogout();
                setIsMenuOpen(false);
              }}
              className="w-full flex items-center gap-4 p-4 rounded-2xl text-red-500 hover:bg-red-50 transition-all font-bold text-sm"
            >
              <LogOut size={20} />
              Đăng xuất
            </button>
          </div>
        </motion.div>
      </>
    )}
  </AnimatePresence>
);

const OfflineBanner = ({ isOnline }: { isOnline: boolean }) => (
  <AnimatePresence>
    {!isOnline && (
      <motion.div 
        initial={{ y: -50 }}
        animate={{ y: 0 }}
        exit={{ y: -50 }}
        className="fixed top-0 left-0 right-0 z-[110] bg-orange-500 text-white py-2 px-4 flex items-center justify-center gap-2 text-[10px] font-bold uppercase tracking-widest"
      >
        <WifiOff size={12} />
        Đang hoạt động ngoại tuyến
      </motion.div>
    )}
  </AnimatePresence>
);

const LoadingOverlay = ({ isSyncing }: { isSyncing: boolean }) => (
  <AnimatePresence>
    {isSyncing && (
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        className="fixed bottom-6 right-6 z-[100] bg-white/80 backdrop-blur-xl p-4 rounded-2xl shadow-2xl border border-black/5 flex items-center gap-3"
      >
        <IOSSpinner />
        <span className="text-xs font-bold text-apple-gray uppercase tracking-widest">Đang đồng bộ...</span>
      </motion.div>
    )}
  </AnimatePresence>
);

export default function App() {
  const [isAuthReady, setIsAuthReady] = useState(false);
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [firebaseUser, setFirebaseUser] = useState<any>(null);

  const pushNotifications = usePushNotifications(firebaseUser?.uid);

  const [currentScreen, setCurrentScreen] = useState<Screen>('login');
  const [screenHistory, setScreenHistory] = useState<Screen[]>(['login']);
  const [loginUsername, setLoginUsername] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  const [isOnline, setIsOnline] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isAddingCategory, setIsAddingCategory] = useState(false);
  const [newCategoryName, setNewCategoryName] = useState('');
  const [isAddingUnit, setIsAddingUnit] = useState(false);
  const [newUnitName, setNewUnitName] = useState('');

  const handleRefresh = async () => {
    if (!navigator.onLine) {
      alert('Không có kết nối mạng. Bạn đang dùng dữ liệu ngoại tuyến.');
      return;
    }
    try {
      setIsSyncing(true);
      // Đảm bảo kết nối mạng Firebase được khôi phục
      await enableNetwork(db);
      // Đợi đến khi tất cả các dữ liệu đã ghi ở chế độ offline được đồng bộ lên máy chủ
      await waitForPendingWrites(db);
      
      // Force refresh data
      await Promise.all([
        getDocs(collection(db, 'inventory')),
        getDocs(collection(db, 'inbound_logs')),
        getDocs(collection(db, 'outbound_logs')),
        getDocs(collection(db, 'audit_logs')),
        getDocs(collection(db, 'settings'))
      ]);
      console.log('Đã làm mới dữ liệu và đồng bộ');
    } catch (error) {
      console.error('Lỗi khi làm mới:', error);
    } finally {
      setIsSyncing(false);
    }
  };

  const navigateTo = (screen: Screen, itemId: string | null = null) => {
    if (itemId) setSelectedItemId(itemId);
    setScreenHistory(prev => [...prev, screen]);
    setCurrentScreen(screen);
    window.scrollTo(0, 0);
  };

  const notifyManagers = async (title: string, body: string, url: string = '/') => {
    try {
      const usersSnapshot = await getDocs(collection(db, 'users'));
      const subscriptions: any[] = [];
      usersSnapshot.forEach(docSnap => {
        const data = docSnap.data();
        if (data.role === 'quản lí' && data.pushSubscriptions && Array.isArray(data.pushSubscriptions)) {
          subscriptions.push(...data.pushSubscriptions);
        }
      });
      
      if (subscriptions.length > 0) {
        await fetch('/api/push/send', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            subscriptions,
            payload: { title, body, url }
          })
        });
      }
    } catch (e) {
      console.error('Không thể gửi push notification:', e);
    }
  };

  const goBack = () => {
    if (screenHistory.length > 1) {
      const newHistory = [...screenHistory];
      newHistory.pop(); // Remove current screen
      const prevScreen = newHistory[newHistory.length - 1];
      setScreenHistory(newHistory);
      setCurrentScreen(prevScreen);
      window.scrollTo(0, 0);
    } else {
      const defaultScreen = currentUser?.role === 'nhân viên' ? 'audit' : 'overview';
      setCurrentScreen(defaultScreen);
      setScreenHistory([defaultScreen]);
    }
  };

  // Components - Moved outside


  const [confirmModal, setConfirmModal] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    onConfirm: () => void;
  }>({
    isOpen: false,
    title: '',
    message: '',
    onConfirm: () => {}
  });
  const [auditType, setAuditType] = useState<AuditType>('daily');
  const [inboundTab, setInboundTab] = useState<'record' | 'history'>('record');
  const [outboundTab, setOutboundTab] = useState<'record' | 'history'>('record');

  // Offline Status Listener
  useEffect(() => {
    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);
    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);
    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
    };
  }, []);

  // Audit Logging Helper
  const logAudit = async (log: LogEntry) => {
    try {
      const logRef = doc(collection(db, 'audit_logs'));
      await setDoc(logRef, {
        ...log,
        timestamp: formatDate(new Date()),
        user: currentUser?.name || 'Unknown'
      });
    } catch (error) {
      console.error('Failed to log audit:', error);
    }
  };
  const [auditCategory, setAuditCategory] = useState<string>('Tất cả');
  const [isEditingAudit, setIsEditingAudit] = useState(false);
  const [selectedItemId, setSelectedItemId] = useState<string | null>(null);
  const [items, setItems] = useState<InventoryItem[]>([]);
  const [itemTransactions, setItemTransactions] = useState<LogEntry[]>([]);
  const [showConfirmModal, setShowConfirmModal] = useState(false);
  const [confirmConfig, setConfirmConfig] = useState<{
    title: string;
    message: string;
    onConfirm: () => void;
    icon?: React.ReactNode;
    type?: 'danger' | 'info';
  } | null>(null);
  const [showSupplierModal, setShowSupplierModal] = useState(false);
  const [editingSupplier, setEditingSupplier] = useState<{
    name: string;
    contact: string;
    price: number;
    idx?: number;
  }>({ name: '', contact: '', price: 0 });
  const [showDeletePasswordModal, setShowDeletePasswordModal] = useState(false);
  const [deletePasswordInput, setDeletePasswordInput] = useState('');
  const [passwordError, setPasswordError] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [inventoryCategory, setInventoryCategory] = useState<string>('Tất cả');
  const [inventoryQuickFilter, setInventoryQuickFilter] = useState<'all' | 'low_stock' | 'urgent' | 'expiring'>('all');
  const [auditSearchQuery, setAuditSearchQuery] = useState('');
  const [inboundSearchQuery, setInboundSearchQuery] = useState('');
  const [outboundSearchQuery, setOutboundSearchQuery] = useState('');
  const [auditInputs, setAuditInputs] = useState<Record<string, number>>({});
  const [purchaseCart, setPurchaseCart] = useState<{ itemId: string; amount: number; price: number }[]>([]);
  const [purchaseInput, setPurchaseInput] = useState<Record<string, string>>({});
  const [globalThreshold, setGlobalThreshold] = useState(10);
  const [customCategories, setCustomCategories] = useState(['Thảo mộc thô', 'Thảo mộc khô', 'Pha chế', 'Vật tư', 'Vệ sinh']);
  const [customUnits, setCustomUnits] = useState(['Gram', 'Kg', 'Cái', 'ml', 'Lưới', 'Thùng', 'Chai', 'Lon', 'Can']);
  const DEFAULT_GAS_WEBHOOK = 'https://script.google.com/macros/s/AKfycbzR3xHMiELN_tDPTcADBMywYaYuRK43WsqMtdM2h_de6iU7p8ZMtXiDKTC3e5D1UgWC/exec';
  const [gasWebhookUrl, setGasWebhookUrl] = useState(DEFAULT_GAS_WEBHOOK);
  const [newItem, setNewItem] = useState<Partial<InventoryItem>>({
    name: '',
    category: 'Thảo mộc thô',
    unit: 'Gram',
    actualStock: 0,
    minThreshold: 0,
    maxThreshold: 0,
    description: '',
    location: '',
    importPrice: 0,
  });
  const [editingItem, setEditingItem] = useState<Partial<InventoryItem> | null>(null);

  const [inboundCart, setInboundCart] = useState<{ itemId: string; amount: number; reason: string }[]>(() => {
    const saved = localStorage.getItem('inventory_app_cache');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.inboundCart || [];
      } catch (e) {
        return [];
      }
    }
    return [];
  });

  const [outboundCart, setOutboundCart] = useState<{ itemId: string; amount: number; reason: string }[]>(() => {
    const saved = localStorage.getItem('inventory_app_cache');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.outboundCart || [];
      } catch (e) {
        return [];
      }
    }
    return [];
  });

  const [currentInbound, setCurrentInbound] = useState({
    itemId: '',
    amount: 0,
    reason: 'Nhập hàng định kỳ',
  });

  const [currentOutbound, setCurrentOutbound] = useState({
    itemId: '',
    amount: 0,
    reason: 'Xuất hàng bán lẻ',
  });

  const [currentTime, setCurrentTime] = useState(new Date());

  useEffect(() => {
    const timer = setInterval(() => {
      setCurrentTime(new Date());
    }, 1000);
    return () => clearInterval(timer);
  }, []);

  const [allLogs, setAllLogs] = useState<LogEntry[]>([]);
  const [reportStartDate, setReportStartDate] = useState(new Date().toISOString().split('T')[0]);
  const [reportEndDate, setReportEndDate] = useState(new Date().toISOString().split('T')[0]);
  const [activeReportStartDate, setActiveReportStartDate] = useState(new Date().toISOString().split('T')[0]);
  const [activeReportEndDate, setActiveReportEndDate] = useState(new Date().toISOString().split('T')[0]);
  const [collapsedReportGroups, setCollapsedReportGroups] = useState<Record<string, boolean>>({});

  const [inboundStartDate, setInboundStartDate] = useState(new Date(new Date().setDate(new Date().getDate() - 30)).toISOString().split('T')[0]);
  const [inboundEndDate, setInboundEndDate] = useState(new Date().toISOString().split('T')[0]);
  const [outboundStartDate, setOutboundStartDate] = useState(new Date(new Date().setDate(new Date().getDate() - 30)).toISOString().split('T')[0]);
  const [outboundEndDate, setOutboundEndDate] = useState(new Date().toISOString().split('T')[0]);
  
  const [editingLog, setEditingLog] = useState<LogEntry | null>(null);
  const [editLogAmount, setEditLogAmount] = useState<string>('');
  const [editLogReason, setEditLogReason] = useState<string>('');

  const [financialStartDate, setFinancialStartDate] = useState(new Date(new Date().setDate(new Date().getDate() - 30)).toISOString().split('T')[0]);
  const [financialEndDate, setFinancialEndDate] = useState(new Date().toISOString().split('T')[0]);
  const [activeFinancialStartDate, setActiveFinancialStartDate] = useState(financialStartDate);
  const [activeFinancialEndDate, setActiveFinancialEndDate] = useState(financialEndDate);

  // Export to Excel
  const exportToExcel = () => {
    const data = items.map(item => ({
      'Tên': item.name,
      'Danh mục': item.category,
      'Đơn vị': item.unit,
      'Tồn kho thực tế': item.actualStock,
      'Vị trí': item.location || 'N/A',
      'Ngưỡng tối thiểu': item.minThreshold || 0,
      'Ngưỡng tối đa': item.maxThreshold || 0,
      'Giá nhập': item.importPrice || 0,
      'Mô tả': item.description || ''
    }));

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Inventory");
    XLSX.writeFile(workbook, `Inventory_Backup_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // Tích hợp Sao lưu và Đồng bộ
  const syncAndBackupData = async () => {
    setIsSyncing(true);
    let successMessages: string[] = [];
    let errorMessages: string[] = [];

    if (!navigator.onLine) {
      errorMessages.push('✗ Không có kết nối mạng. Đồng bộ thất bại.');
    } else {
      try {
        await enableNetwork(db);
        await waitForPendingWrites(db);
        successMessages.push('✓ Dữ liệu cục bộ đã đồng bộ với Cloud');
      } catch (error: any) {
        errorMessages.push(`✗ Lỗi kết nối Cloud (${error.message})`);
      }

      // 1. Sao lưu Cloud (Firestore)
      try {
        const backupData = {
          timestamp: new Date().toISOString(),
          inventory: items,
          settings: {
            globalThreshold,
            customCategories,
            customUnits
          },
          logs: allLogs,
          backupBy: currentUser?.username || 'System'
        };
        await setDoc(doc(db, 'backups', `backup_${Date.now()}`), backupData);
        successMessages.push('✓ Firestore Cloud (Tạo bản sao lưu thành công)');
      } catch (error: any) {
        errorMessages.push(`✗ Firestore Cloud (${error.message || 'Lỗi không xác định'})`);
      }

      // 2. Apps Script (Webhook)
      const webhookUrl = gasWebhookUrl || DEFAULT_GAS_WEBHOOK;
      if (webhookUrl) {
        try {
          const payload = {
            action: 'sync_inventory',
            sheetId: '1P7PGYjDZFxIT1kG_EmuFI91omVHKBFHqLMgi-GtAbZU',
            timestamp: new Date().toISOString(),
            inventory: items.map(item => ({
              id: item.id,
              name: item.name,
              category: item.category,
              unit: item.unit,
              actualStock: item.actualStock,
              location: item.location || '',
              minThreshold: item.minThreshold || 0,
              importPrice: item.importPrice || 0,
              description: item.description || ''
            })),
            inboundLogs: allLogs.filter(l => l.type === 'Nhập kho'),
            outboundLogs: allLogs.filter(l => l.type === 'Xuất kho'),
            auditLogs: allLogs.filter(l => l.type === 'Kiểm kê')
          };

          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 15000); // 15 seconds timeout

          await fetch(webhookUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(payload),
            redirect: 'follow',
            signal: controller.signal
          });
          clearTimeout(timeoutId);
          successMessages.push('✓ Apps Script Webhook (Đã gửi)');
        } catch (error: any) {
          if (error.name === 'AbortError') {
             errorMessages.push('✗ Apps Script Webhook (Hết thời gian phản hồi)');
          } else {
             errorMessages.push(`✗ Apps Script Webhook (${error.message})`);
          }
        }
      } else {
        errorMessages.push('✗ Apps Script Webhook (Thiếu URL cấu hình)');
      }
    }

    setIsSyncing(false);

    if (errorMessages.length === 0) {
      setConfirmConfig({
        title: 'Đồng bộ hoàn tất',
        message: 'Tất cả các dịch vụ đã được cập nhật thành công:\n\n' + successMessages.join('\n'),
        onConfirm: () => setShowConfirmModal(false),
        icon: <Cloud className="w-10 h-10 text-apple-green" />,
        type: 'info'
      });
    } else {
      setConfirmConfig({
        title: 'Hoàn tất với một số lỗi',
        message: `Thành công:\n${successMessages.join('\n')}\n\nLỗi:\n${errorMessages.join('\n')}`,
        onConfirm: () => setShowConfirmModal(false),
        icon: <AlertTriangle className="w-10 h-10 text-orange-500" />,
        type: 'danger'
      });
    }
    setShowConfirmModal(true);
  };

  // Auth and Real-time Sync
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      setFirebaseUser(user);
      const savedUser = localStorage.getItem('botanical_user');
      if (savedUser) {
        try {
          const userData = JSON.parse(savedUser);
          setCurrentUser(userData);
          
          // If signed in to app but not to Firebase, sign in anonymously
          if (!user) {
            try {
              await signInAnonymously(auth);
            } catch (authErr: any) {
              if (authErr.code === 'auth/admin-restricted-operation') {
                console.error('LỖI QUAN TRỌNG: Bạn cần bật "Anonymous Authentication" trong Firebase Console (Authentication -> Sign-in method -> Anonymous -> Enable).');
                setLoginError('Hệ thống chưa được cấu hình đầy đủ. Vui lòng liên hệ quản trị viên để bật Anonymous Auth trong Firebase.');
              } else {
                throw authErr;
              }
            }
            // The next onAuthStateChanged call will handle the rest
            return;
          }

          // Sync user profile to Firestore so rules can check role
          if (user) {
            await setDoc(doc(db, 'users', user.uid), {
              uid: user.uid,
              username: userData.username,
              role: userData.role,
              name: userData.name
            }, { merge: true });
          }

          if (currentScreen === 'login') {
            setCurrentScreen(userData.role === 'quản lí' ? 'overview' : 'audit');
          }
        } catch (e) {
          console.error('Auth initialization error:', e);
          localStorage.removeItem('botanical_user');
        }
      }
      setIsAuthReady(true);
    });
    return () => unsubscribe();
  }, []);

  // Sync Inventory Items
  useEffect(() => {
    if (!currentUser || !firebaseUser) return;

    const unsubscribeItems = onSnapshot(collection(db, 'inventory'), (snapshot) => {
      const newItems = snapshot.docs.map(doc => ({ ...doc.data() as InventoryItem, id: doc.id }));
      setItems(newItems);
    }, (error) => handleFirestoreError(error, OperationType.GET, 'inventory'));

    return () => unsubscribeItems();
  }, [currentUser, firebaseUser]);

  useEffect(() => {
    if (customCategories.length > 0 && !customCategories.includes(newItem.category as string)) {
      setNewItem(prev => ({ ...prev, category: customCategories[0] }));
    }
    if (customUnits.length > 0 && !customUnits.includes(newItem.unit as string)) {
      setNewItem(prev => ({ ...prev, unit: customUnits[0] }));
    }
  }, [customCategories, customUnits]);

  useEffect(() => {
    if (currentScreen === 'edit' && selectedItem) {
      setEditingItem({ ...selectedItem });
    } else if (currentScreen !== 'edit') {
      setEditingItem(null);
    }
  }, [currentScreen, selectedItemId, items]);

  useEffect(() => {
    if (editingItem && customCategories.length > 0 && !customCategories.includes(editingItem.category as string)) {
      setEditingItem(prev => prev ? ({ ...prev, category: customCategories[0] }) : null);
    }
    if (editingItem && customUnits.length > 0 && !customUnits.includes(editingItem.unit as string)) {
      setEditingItem(prev => prev ? ({ ...prev, unit: customUnits[0] }) : null);
    }
  }, [customCategories, customUnits, editingItem?.id]);

  // Sync Settings
  useEffect(() => {
    if (!currentUser || !firebaseUser) return;

    const unsubscribeSettings = onSnapshot(doc(db, 'settings', 'global'), (snapshot) => {
      if (snapshot.exists()) {
        const data = snapshot.data();
        if (data.globalThreshold !== undefined) setGlobalThreshold(data.globalThreshold);
        if (data.customCategories) setCustomCategories(data.customCategories);
        if (data.customUnits) setCustomUnits(data.customUnits);
        if (data.gasWebhookUrl) setGasWebhookUrl(data.gasWebhookUrl);
      }
    }, (error) => handleFirestoreError(error, OperationType.GET, 'settings/global'));

    // Sync User Specific Settings
    const unsubscribeUserSettings = onSnapshot(doc(db, 'users', firebaseUser.uid), (snapshot) => {
      if (snapshot.exists()) {
        // user settings can be handled here if added in the future
      }
    }, (error) => handleFirestoreError(error, OperationType.GET, `users/${firebaseUser.uid}`));

    return () => {
      unsubscribeSettings();
      unsubscribeUserSettings();
    };
  }, [currentUser, firebaseUser]);

  // Sync All Logs for Reports/Financials
  useEffect(() => {
    if (!currentUser || !firebaseUser) return;

    const unsubInbound = onSnapshot(collection(db, 'inbound_logs'), (snapshot) => {
      const logs = snapshot.docs.map(doc => ({ ...doc.data(), docId: doc.id, type: 'Nhập kho' } as LogEntry));
      setAllLogs(prev => {
        const otherLogs = prev.filter(l => l.type !== 'Nhập kho');
        return [...otherLogs, ...logs].sort((a, b) => b.timestamp.localeCompare(a.timestamp));
      });
    }, (error) => handleFirestoreError(error, OperationType.GET, 'inbound_logs'));

    const unsubOutbound = onSnapshot(collection(db, 'outbound_logs'), (snapshot) => {
      const logs = snapshot.docs.map(doc => ({ ...doc.data(), docId: doc.id, type: 'Xuất kho' } as LogEntry));
      setAllLogs(prev => {
        const otherLogs = prev.filter(l => l.type !== 'Xuất kho');
        return [...otherLogs, ...logs].sort((a, b) => b.timestamp.localeCompare(a.timestamp));
      });
    }, (error) => handleFirestoreError(error, OperationType.GET, 'outbound_logs'));

    const unsubAudit = onSnapshot(collection(db, 'audit_logs'), (snapshot) => {
      const logs = snapshot.docs.map(doc => ({ ...doc.data(), docId: doc.id, type: 'Kiểm kê' } as LogEntry));
      setAllLogs(prev => {
        const otherLogs = prev.filter(l => l.type !== 'Kiểm kê');
        return [...otherLogs, ...logs].sort((a, b) => b.timestamp.localeCompare(a.timestamp));
      });
    }, (error) => handleFirestoreError(error, OperationType.GET, 'audit_logs'));

    return () => {
      unsubInbound();
      unsubOutbound();
      unsubAudit();
    };
  }, [currentUser, firebaseUser]);

  // Test Connection
  useEffect(() => {
    async function testConnection() {
      try {
        await getDocFromServer(doc(db, 'test', 'connection'));
      } catch (error) {
        if (error instanceof Error && error.message.includes('the client is offline')) {
          console.error("Please check your Firebase configuration.");
        }
      }
    }
    testConnection();
  }, []);

  const addToInboundCart = () => {
    if (!currentInbound.itemId || currentInbound.amount <= 0) return;
    setInboundCart(prev => [...prev, { ...currentInbound }]);
    setCurrentInbound({
      itemId: '',
      amount: 0,
      reason: 'Nhập hàng định kỳ',
    });
  };

  const removeFromInboundCart = (index: number) => {
    setInboundCart(prev => prev.filter((_, i) => i !== index));
  };

  const addToOutboundCart = () => {
    if (!currentOutbound.itemId || currentOutbound.amount <= 0) return;
    setOutboundCart(prev => [...prev, { ...currentOutbound }]);
    setCurrentOutbound({
      itemId: '',
      amount: 0,
      reason: 'Xuất hàng bán lẻ',
    });
  };

  const removeFromOutboundCart = (index: number) => {
    setOutboundCart(prev => prev.filter((_, i) => i !== index));
  };

  const handleDeleteLog = async (log: LogEntry) => {
    if (!log.docId || !log.id || !log.amount) return;
    
    setConfirmModal({
      isOpen: true,
      title: 'Xóa lịch sử',
      message: `Bạn có chắc chắn muốn xóa lịch sử ${log.type.toLowerCase()} này? Số lượng tồn kho sẽ được hoàn tác.`,
      onConfirm: async () => {
        try {
          setIsSyncing(true);
          const batch = writeBatch(db);
          
          // Revert inventory
          const itemRef = doc(db, 'inventory', log.id);
          if (log.amount !== undefined) {
            const stockChange = log.type === 'Nhập kho' ? -log.amount : log.amount;
            batch.update(itemRef, { actualStock: increment(stockChange) });
          }

          // Delete log
          const collectionName = log.type === 'Nhập kho' ? 'inbound_logs' : 'outbound_logs';
          const logRef = doc(db, collectionName, log.docId!);
          batch.delete(logRef);

          // Add audit log
          const auditLogRef = doc(collection(db, 'audit_logs'));
          batch.set(auditLogRef, {
            id: log.id,
            name: log.name,
            type: 'Xóa lịch sử',
            timestamp: formatDate(new Date()),
            user: currentUser?.name || 'Unknown',
            details: `Đã xóa lịch sử ${log.type.toLowerCase()} (${log.amount} ${log.unit || ''})`
          });

          await batch.commit();
        } catch (error) {
          handleFirestoreError(error, OperationType.DELETE, 'logs');
        } finally {
          setIsSyncing(false);
          setConfirmModal(prev => ({ ...prev, isOpen: false }));
        }
      }
    });
  };

  const handleEditLogSubmit = async () => {
    if (!editingLog || !editingLog.docId || !editingLog.id || !editingLog.amount) return;
    const newAmount = parseFloat(editLogAmount);
    if (isNaN(newAmount) || newAmount <= 0) return;

    try {
      setIsSyncing(true);
      const batch = writeBatch(db);
      
      // Update inventory
      const itemRef = doc(db, 'inventory', editingLog.id);
      let stockChange = 0;
      if (editingLog.type === 'Nhập kho') {
        stockChange = newAmount - editingLog.amount;
      } else if (editingLog.type === 'Xuất kho') {
        stockChange = editingLog.amount - newAmount;
      }
      if (stockChange !== 0) {
        batch.update(itemRef, { actualStock: increment(stockChange) });
      }

      // Update log
      const collectionName = editingLog.type === 'Nhập kho' ? 'inbound_logs' : 'outbound_logs';
      const logRef = doc(db, collectionName, editingLog.docId);
      batch.update(logRef, {
        amount: newAmount,
        reason: editLogReason || editingLog.reason,
      });

      // Add audit log
      const auditLogRef = doc(collection(db, 'audit_logs'));
      batch.set(auditLogRef, {
        id: editingLog.id,
        name: editingLog.name,
        type: 'Sửa lịch sử',
        timestamp: formatDate(new Date()),
        user: currentUser?.name || 'Unknown',
        details: `Đã sửa lịch sử ${editingLog.type.toLowerCase()} từ ${editingLog.amount} thành ${newAmount} ${editingLog.unit || ''}`
      });

      await batch.commit();
      setEditingLog(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, 'logs');
    } finally {
      setIsSyncing(false);
    }
  };

  const handleInboundSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (inboundCart.length === 0) return;

    const now = formatDate(new Date());
    const syncData = inboundCart.map(cartItem => {
      const item = items.find(i => i.id === cartItem.itemId);
      if (!item) return null;
      const newStock = (item.actualStock || 0) + cartItem.amount;
      const log: LogEntry = {
        id: item.id,
        name: item.name,
        type: 'Nhập kho',
        amount: cartItem.amount,
        oldStock: item.actualStock || 0,
        newStock: newStock,
        reason: cartItem.reason,
        timestamp: now,
        user: currentUser?.name || 'Unknown',
        importPrice: item.importPrice || 0,
        unit: item.unit || '',
      };
      return log;
    }).filter((log): log is LogEntry => log !== null);

    try {
      setIsSyncing(true);
      const batch = writeBatch(db);
      
      // Update inventory items
      inboundCart.forEach(cartItem => {
        const itemRef = doc(db, 'inventory', cartItem.itemId);
        batch.update(itemRef, { actualStock: increment(cartItem.amount) });
      });

      // Add logs
      syncData.forEach(log => {
        const logRef = doc(collection(db, 'inbound_logs'));
        batch.set(logRef, log);
        
        // Also log to audit_logs for unified tracking
        const auditLogRef = doc(collection(db, 'audit_logs'));
        batch.set(auditLogRef, {
          ...log,
          details: `Nhập kho thêm ${log.amount} ${log.unit} (${log.reason})`
        });
      });

      await batch.commit();

      // Background auto-sync to Sheets
      const updatedItemsForSync = items.map(initialItem => {
        const cartMatch = inboundCart.find(c => c.itemId === initialItem.id);
        if (cartMatch) {
          return { ...initialItem, actualStock: (initialItem.actualStock || 0) + cartMatch.amount };
        }
        return initialItem;
      });
      fetch('/api/sync/sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ inventory: updatedItemsForSync })
      }).catch(err => console.error('Background sync failed:', err));

      if (gasWebhookUrl) {
        const logsWithCategory = syncData.map(log => {
          const item = items.find(i => i.id === log.id);
          return {
            ...log,
            category: item?.category || '',
            price: item?.importPrice,
            stockAfter: log.newStock
          };
        });

        fetch(gasWebhookUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain' },
          body: JSON.stringify({
            action: 'log_inventory_change',
            logs: logsWithCategory
          })
        }).catch(err => console.error('Background GAS sync failed:', err));
      }

      setInboundCart([]);
      navigateTo('overview');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'inbound');
    } finally {
      setIsSyncing(false);
    }
  };

  const handleOutboundSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (outboundCart.length === 0) return;

    const now = formatDate(new Date());
    const syncData = outboundCart.map(cartItem => {
      const item = items.find(i => i.id === cartItem.itemId);
      if (!item) return null;
      const newStock = (item.actualStock || 0) - cartItem.amount;
      const log: LogEntry = {
        id: item.id,
        name: item.name,
        type: 'Xuất kho',
        amount: cartItem.amount,
        oldStock: item.actualStock || 0,
        newStock: newStock,
        reason: cartItem.reason,
        timestamp: now,
        user: currentUser?.name || 'Unknown',
        importPrice: item.importPrice || 0,
        unit: item.unit || '',
      };
      return log;
    }).filter((log): log is LogEntry => log !== null);

    try {
      setIsSyncing(true);
      const batch = writeBatch(db);
      
      // Update inventory items
      outboundCart.forEach(cartItem => {
        const itemRef = doc(db, 'inventory', cartItem.itemId);
        batch.update(itemRef, { actualStock: increment(-cartItem.amount) });
      });

      // Add logs
      syncData.forEach(log => {
        const logRef = doc(collection(db, 'outbound_logs'));
        batch.set(logRef, log);
        
        // Also log to audit_logs for unified tracking
        const auditLogRef = doc(collection(db, 'audit_logs'));
        batch.set(auditLogRef, {
          ...log,
          details: `Xuất kho ${log.amount} ${log.unit} (${log.reason})`
        });
      });

      await batch.commit();

      // Background auto-sync to Sheets
      const updatedItemsForSync = items.map(initialItem => {
        const cartMatch = outboundCart.find(c => c.itemId === initialItem.id);
        if (cartMatch) {
          return { ...initialItem, actualStock: (initialItem.actualStock || 0) - cartMatch.amount };
        }
        return initialItem;
      });
      fetch('/api/sync/sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ inventory: updatedItemsForSync })
      }).catch(err => console.error('Background sync failed:', err));

      if (gasWebhookUrl) {
        const logsWithCategory = syncData.map(log => {
          const item = items.find(i => i.id === log.id);
          return {
            ...log,
            category: item?.category || '',
            price: item?.importPrice,
            stockAfter: log.newStock
          };
        });

        fetch(gasWebhookUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain' },
          body: JSON.stringify({
            action: 'log_inventory_change',
            logs: logsWithCategory
          })
        }).catch(err => console.error('Background GAS sync failed:', err));
      }

      // Check for low stock after outbound
      const lowStockItems = updatedItemsForSync.filter(item => {
        const threshold = item.minThreshold || globalThreshold;
        return item.actualStock > 0 && item.actualStock <= threshold;
      });
      // Filter out items that were already low stock BEFORE this transaction (to avoid spam), 
      // by checking if they just dropped below threshold in this transaction
      const newlyLowStockItems = lowStockItems.filter(item => {
        const cartMatch = outboundCart.find(c => c.itemId === item.id);
        if (cartMatch) {
          const oldStock = item.actualStock + cartMatch.amount;
          const threshold = item.minThreshold || globalThreshold;
          return oldStock > threshold && item.actualStock <= threshold;
        }
        return false;
      });

      if (newlyLowStockItems.length > 0) {
        notifyManagers(
          'Cảnh báo Tồn kho',
          `Có ${newlyLowStockItems.length} mặt hàng vừa tụt xuống dưới mức cảnh báo (${newlyLowStockItems.map(i => i.name).join(', ')}).`,
          '/suppliers'
        );
      }

      setOutboundCart([]);
      navigateTo('overview');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'outbound');
    } finally {
      setIsSyncing(false);
    }
  };

  const handleAuditConfirm = async () => {
    let auditItemsToSync: InventoryItem[] = [];
    
    if (currentScreen === 'details' && selectedItem) {
      const freq = selectedItem.auditFrequency?.[0] || 'daily';
      if (freq === 'daily') {
        auditItemsToSync = items.filter(i => i.auditFrequency?.includes('daily'));
      } else {
        auditItemsToSync = [selectedItem];
      }
    } else {
      auditItemsToSync = items.filter(i => i.auditFrequency?.includes(auditType));
    }

    const now = formatDate(new Date());
    let typeToReport = auditType;
    
    if (currentScreen === 'details' && selectedItem) {
      typeToReport = selectedItem.auditFrequency?.[0] || 'daily';
    }

    const auditData: LogEntry[] = auditItemsToSync.map(item => {
      const actual = auditInputs[item.id] !== undefined ? auditInputs[item.id] : item.actualStock;
      const diff = actual - item.actualStock;
      
      return {
        id: item.id,
        name: item.name,
        category: item.category,
        theoreticalStock: item.actualStock,
        actualStock: actual,
        unit: item.unit,
        type: 'Kiểm kê',
        timestamp: now,
        user: currentUser?.name || 'Unknown',
        reason: diff === 0 ? 'Khớp tồn kho' : (diff > 0 ? 'Thừa hàng' : 'Thiếu hàng'),
        oldStock: item.actualStock,
        newStock: actual,
        details: diff !== 0 ? `Điều chỉnh tồn kho từ ${item.actualStock} sang ${actual} (${diff > 0 ? '+' : ''}${diff})` : 'Tồn kho khớp'
      };
    });

    try {
      setIsSyncing(true);
      const batch = writeBatch(db);
      
      auditData.forEach(log => {
        const itemRef = doc(db, 'inventory', log.id);
        batch.update(itemRef, { actualStock: log.actualStock });
        
        const logRef = doc(collection(db, 'audit_logs'));
        batch.set(logRef, log);
      });

      await batch.commit();

      // Background auto-sync to Sheets
      const updatedItemsForSync = items.map(initialItem => {
        const auditMatch = auditData.find(a => a.id === initialItem.id);
        if (auditMatch) {
          return { ...initialItem, actualStock: auditMatch.actualStock };
        }
        return initialItem;
      });
      fetch('/api/sync/sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ inventory: updatedItemsForSync })
      }).catch(err => console.error('Background sync failed:', err));

      // Send audit report to GAS if configured
      if (gasWebhookUrl) {
        fetch(gasWebhookUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'audit_report',
            auditType: typeToReport,
            auditData: auditData
          })
        }).catch(err => console.error('GAS Audit Report failed:', err));

        const logsWithCategory = auditData.filter(log => (log.actualStock ?? 0) !== (log.theoreticalStock ?? 0)).map(log => {
          const item = items.find(i => i.id === log.id);
          const diff = (log.actualStock ?? 0) - (log.theoreticalStock ?? 0);
          return {
            ...log,
            category: item?.category || '',
            price: item?.importPrice,
            stockAfter: (log.actualStock ?? 0),
            amount: Math.abs(diff),
            type: diff > 0 ? 'Nhập kho (Kiểm kê thừa)' : 'Xuất kho (Kiểm kê thiếu)',
            reason: log.reason
          };
        });

        if (logsWithCategory.length > 0) {
          fetch(gasWebhookUrl, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
              action: 'log_inventory_change',
              logs: logsWithCategory
            })
          }).catch(err => console.error('Background GAS sync failed:', err));
        }
      }

      // Notify managers about audit completion
      notifyManagers(
        'Hoàn thành kiểm kê',
        `${currentUser?.name || 'Ai đó'} vừa hoàn tất kiểm kê ${typeToReport === 'daily' ? 'ngày' : typeToReport === 'weekly' ? 'tuần' : 'tháng'}. Đã kiểm kê ${auditData.length} mặt hàng.`,
        '/report'
      );

      // Check for low stock
      const lowStockItems = updatedItemsForSync.filter(item => 
        (item.actualStock ?? 0) > 0 && (item.actualStock ?? 0) <= (item.minThreshold || ((item.actualStock ?? 0) < globalThreshold ? globalThreshold : 0))
      );
      if (lowStockItems.length > 0) {
        notifyManagers(
          'Cảnh báo Tồn kho Thấp',
          `Có ${lowStockItems.length} mặt hàng sắp hết. Cần đặt hàng thêm!`,
          '/suppliers'
        );
      }

      setAuditInputs({});
      setShowConfirmModal(false);
      setConfirmConfig(null);
      navigateTo('overview');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'audit');
    } finally {
      setIsSyncing(false);
    }
  };

  const handleAddPurchaseCart = (item: InventoryItem) => {
    const amountStr = purchaseInput[item.id];
    const amount = Number(amountStr);
    if (!amount || amount <= 0) return;

    let bestPrice = item.importPrice || 0;
    if (item.suppliers && item.suppliers.length > 0) {
      bestPrice = Math.min(...item.suppliers.map(s => s.price));
    }

    setPurchaseCart(prev => {
      const existing = prev.find(p => p.itemId === item.id);
      if (existing) {
        return prev.map(p => p.itemId === item.id ? { ...p, amount: p.amount + amount } : p);
      }
      return [...prev, { itemId: item.id, amount, price: bestPrice }];
    });
    setPurchaseInput(prev => ({ ...prev, [item.id]: '' }));
  };

  const handleRemovePurchaseCart = (itemId: string) => {
    setPurchaseCart(prev => prev.filter(p => p.itemId !== itemId));
  };

  const handlePreAudit = () => {
    const today = formatDate(new Date()).split(' - ')[0];
    const auditItemsToSync = items.filter(i => 
      i.auditFrequency?.includes(auditType) && 
      (auditCategory === 'Tất cả' || i.category === auditCategory)
    );
    
    const duplicates = allLogs.filter(log => 
      log.type === 'Kiểm kê' && 
      log.timestamp.startsWith(today) &&
      auditItemsToSync.some(i => i.id === log.id)
    ).map(l => l.name);

    if (duplicates.length > 0) {
      setConfirmConfig({
        title: 'Cảnh báo trùng lặp',
        message: `Các dược liệu sau đã được kiểm kê hôm nay: ${duplicates.join(', ')}. Bạn có chắc chắn muốn kiểm kê lại không?`,
        onConfirm: () => {
          handleAuditConfirm();
          setShowConfirmModal(false);
          setConfirmConfig(null);
        },
        type: 'danger',
        icon: <AlertTriangle className="w-10 h-10 text-red-500" />
      });
      setShowConfirmModal(true);
    } else {
      setConfirmConfig({
        title: 'Xác nhận kiểm kê',
        message: `Đồng bộ kết quả kiểm kê cho ${auditItemsToSync.length} dược liệu vào hệ thống?`,
        onConfirm: () => {
          handleAuditConfirm();
          setShowConfirmModal(false);
          setConfirmConfig(null);
        },
        type: 'info',
        icon: <ClipboardList className="w-10 h-10 text-apple-blue" />
      });
      setShowConfirmModal(true);
    }
  };

  const handlePreAddItem = () => {
    if (!newItem.name) return;
    const today = formatDate(new Date()).split(' - ')[0];
    const duplicate = allLogs.find(log => 
      log.type === 'Thêm mới' && 
      log.name.toLowerCase() === newItem.name?.toLowerCase() &&
      log.timestamp.startsWith(today)
    );

    if (duplicate) {
      setConfirmConfig({
        title: 'Cảnh báo trùng lặp',
        message: `Dược liệu "${newItem.name}" đã được thêm mới vào hệ thống hôm nay. Bạn có chắc chắn muốn thêm tiếp không?`,
        onConfirm: () => {
          handleAddItem();
          setShowConfirmModal(false);
          setConfirmConfig(null);
        },
        type: 'danger',
        icon: <AlertTriangle className="w-10 h-10 text-red-500" />
      });
      setShowConfirmModal(true);
    } else {
      handleAddItem();
    }
  };

  const handleDeleteItem = async (itemId: string) => {
    try {
      setIsSyncing(true);
      const item = items.find(i => i.id === itemId);
      if (item) {
        await deleteDoc(doc(db, 'inventory', itemId));
        await logAudit({
          id: itemId,
          name: item.name,
          type: 'Xóa dược liệu',
          amount: 0,
          oldStock: item.actualStock,
          newStock: 0,
          reason: 'Xóa dược liệu khỏi hệ thống',
          timestamp: formatDate(new Date()),
          user: currentUser?.name || 'Unknown',
          details: `Xóa dược liệu ${item.name} khỏi danh mục`
        });

        // Background auto-sync to Sheets
        fetch('/api/sync/sheets', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ inventory: items.filter(i => i.id !== itemId) })
        }).catch(err => console.error('Background sync failed:', err));

        if (gasWebhookUrl) {
          fetch(gasWebhookUrl, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
              action: 'log_inventory_change',
              logs: [{
                user: currentUser?.name || 'Unknown',
                type: 'Xóa dược liệu',
                name: item.name,
                category: item.category,
                amount: item.actualStock,
                unit: item.unit,
                price: item.importPrice || 0,
                reason: 'Hủy dữ liệu, tồn kho về 0',
                stockAfter: 0
              }]
            })
          }).catch(err => console.error('Background GAS sync failed:', err));
        }
      }
      navigateTo('inventory');
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, `inventory/${itemId}`);
    } finally {
      setIsSyncing(false);
      setShowConfirmModal(false);
      setConfirmConfig(null);
    }
  };

  const handleSaveSupplier = async () => {
    if (!selectedItem || !editingSupplier.name) return;

    try {
      setIsSyncing(true);
      const currentSuppliers = selectedItem.suppliers || [];
      let updatedSuppliers = [...currentSuppliers];

      if (editingSupplier.idx !== undefined) {
        updatedSuppliers[editingSupplier.idx] = {
          id: currentSuppliers[editingSupplier.idx].id || Math.random().toString(36).substr(2, 9),
          name: editingSupplier.name,
          contact: editingSupplier.contact,
          price: editingSupplier.price,
          lastUpdated: formatDate(new Date())
        };
      } else {
        updatedSuppliers.push({
          id: Math.random().toString(36).substr(2, 9),
          name: editingSupplier.name,
          contact: editingSupplier.contact,
          price: editingSupplier.price,
          lastUpdated: formatDate(new Date())
        });
      }

      await updateDoc(doc(db, 'inventory', selectedItem.id), {
        suppliers: updatedSuppliers
      });

      setShowSupplierModal(false);
      setEditingSupplier({ name: '', contact: '', price: 0 });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `inventory/${selectedItem.id}/suppliers`);
    } finally {
      setIsSyncing(false);
    }
  };

  const removeAccents = (str: string) => {
    return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  };

  const fuse = useMemo(() => {
    const itemsWithNormalized = items.map(item => ({
      ...item,
      normalizedName: removeAccents(item.name),
      normalizedCategory: removeAccents(item.category),
      normalizedDescription: removeAccents(item.description || ''),
    }));

    return new Fuse(itemsWithNormalized, {
      keys: ['name', 'category', 'description', 'normalizedName', 'normalizedCategory', 'normalizedDescription'],
      threshold: 0.4,
      location: 0,
      distance: 100,
      minMatchCharLength: 1,
      useExtendedSearch: true,
    });
  }, [items]);

  const inventoryHealth = useMemo(() => {
    if (items.length === 0) return 100;
    const healthyItems = items.filter(item => item.minThreshold !== undefined && item.actualStock >= item.minThreshold).length;
    return Math.round((healthyItems / items.length) * 100);
  }, [items]);

  const totalStockWeight = useMemo(() => {
    return items.reduce((acc, item) => acc + item.actualStock, 0).toFixed(1);
  }, [items]);

  const filteredLogsForFinancial = useMemo(() => {
    return allLogs.filter(log => {
      if (!log.timestamp || typeof log.timestamp !== 'string') return false;
      const parts = log.timestamp.split(' - ');
      const datePart = parts[0];
      if (!datePart) return false;
      const dateParts = datePart.split('/');
      if (dateParts.length < 3) return false;
      // Ensure month and day are padded with 0
      const d = dateParts[0].padStart(2, '0');
      const m = dateParts[1].padStart(2, '0');
      const y = dateParts[2];
      const logDate = `${y}-${m}-${d}`;
      return logDate >= activeFinancialStartDate && logDate <= activeFinancialEndDate;
    });
  }, [allLogs, activeFinancialStartDate, activeFinancialEndDate]);

  const totalInboundValueInRange = useMemo(() => {
    return filteredLogsForFinancial
      .filter(log => log.type === 'Nhập kho' && log.amount && log.importPrice)
      .reduce((sum, log) => sum + (log.amount! * log.importPrice!), 0);
  }, [filteredLogsForFinancial]);
  const totalInventoryValue = useMemo(() => {
    return items.reduce((acc, item) => {
      const price = item.importPrice || 0;
      return acc + (item.actualStock * price);
    }, 0);
  }, [items]);

  const searchSuggestions = useMemo(() => {
    if (!searchQuery || searchQuery.length < 2) return [];
    const normalizedQuery = removeAccents(searchQuery);
    const results = fuse.search(normalizedQuery);
    return results.slice(0, 5).map(r => r.item);
  }, [searchQuery, fuse]);

  const filteredItems = useMemo(() => {
    let baseItems = items;
    if (inventoryCategory !== 'Tất cả') {
      baseItems = baseItems.filter(i => i.category === inventoryCategory);
    }

    if (inventoryQuickFilter === 'low_stock') {
      baseItems = baseItems.filter(i => i.actualStock > 0 && i.minThreshold !== undefined && i.actualStock <= i.minThreshold);
    } else if (inventoryQuickFilter === 'urgent') {
      baseItems = baseItems.filter(i => i.actualStock === 0);
    } else if (inventoryQuickFilter === 'expiring') {
      const thirtyDaysFromNow = new Date();
      thirtyDaysFromNow.setDate(thirtyDaysFromNow.getDate() + 30);
      baseItems = baseItems.filter(i => {
        if (!i.expiryDate) return false;
        const expiry = new Date(i.expiryDate);
        return expiry <= thirtyDaysFromNow && i.actualStock > 0;
      });
    }
    
    if (!searchQuery) return baseItems;
    
    const normalizedQuery = removeAccents(searchQuery);
    const results = fuse.search(normalizedQuery);
    
    const matchedItems = results.map(result => result.item);
    
    let filteredMatched = matchedItems;
    if (inventoryCategory !== 'Tất cả') {
      filteredMatched = filteredMatched.filter(i => i.category === inventoryCategory);
    }
    if (inventoryQuickFilter === 'low_stock') {
      filteredMatched = filteredMatched.filter(i => i.actualStock > 0 && i.minThreshold !== undefined && i.actualStock <= i.minThreshold);
    } else if (inventoryQuickFilter === 'urgent') {
      filteredMatched = filteredMatched.filter(i => i.actualStock === 0);
    } else if (inventoryQuickFilter === 'expiring') {
      const thirtyDaysFromNow = new Date();
      thirtyDaysFromNow.setDate(thirtyDaysFromNow.getDate() + 30);
      filteredMatched = filteredMatched.filter(i => {
        if (!i.expiryDate) return false;
        const expiry = new Date(i.expiryDate);
        return expiry <= thirtyDaysFromNow && i.actualStock > 0;
      });
    }
    
    return filteredMatched;
  }, [searchQuery, items, fuse, inventoryCategory, inventoryQuickFilter]);

  const filteredAuditItems = useMemo(() => {
    const baseItems = isEditingAudit ? items : items.filter(i => i.auditFrequency?.includes(auditType));
    const categoryFiltered = baseItems.filter(i => auditCategory === 'Tất cả' || i.category === auditCategory);
    
    if (!auditSearchQuery) return categoryFiltered;
    
    const normalizedQuery = removeAccents(auditSearchQuery);
    const auditFuse = new Fuse(categoryFiltered.map(item => ({
      ...item,
      normalizedName: removeAccents(item.name),
    })), {
      keys: ['name', 'normalizedName'],
      threshold: 0.3,
    });
    
    return auditFuse.search(normalizedQuery).map((result: any) => result.item);
  }, [auditSearchQuery, items, auditType, auditCategory, isEditingAudit]);

  const auditSearchSuggestions = useMemo(() => {
    if (!auditSearchQuery || auditSearchQuery.length < 2) return [];
    const normalizedQuery = removeAccents(auditSearchQuery);
    const auditFuse = new Fuse(items.map(item => ({
      ...item,
      normalizedName: removeAccents(item.name),
    })), {
      keys: ['name', 'normalizedName'],
      threshold: 0.3,
    });
    const results = auditFuse.search(normalizedQuery);
    return results.slice(0, 5).map((r: any) => r.item);
  }, [auditSearchQuery, items]);

  const inboundSearchSuggestions = useMemo(() => {
    if (!inboundSearchQuery || inboundSearchQuery.length < 2) return [];
    const normalizedQuery = removeAccents(inboundSearchQuery);
    const inboundFuse = new Fuse(items.map(item => ({
      ...item,
      normalizedName: removeAccents(item.name),
    })), {
      keys: ['name', 'normalizedName'],
      threshold: 0.3,
    });
    const results = inboundFuse.search(normalizedQuery);
    return results.slice(0, 5).map((r: any) => r.item);
  }, [inboundSearchQuery, items]);

  const outboundSearchSuggestions = useMemo(() => {
    if (!outboundSearchQuery || outboundSearchQuery.length < 2) return [];
    const normalizedQuery = removeAccents(outboundSearchQuery);
    const outboundFuse = new Fuse(items.map(item => ({
      ...item,
      normalizedName: removeAccents(item.name),
    })), {
      keys: ['name', 'normalizedName'],
      threshold: 0.3,
    });
    const results = outboundFuse.search(normalizedQuery);
    return results.slice(0, 5).map((r: any) => r.item);
  }, [outboundSearchQuery, items]);

  const selectedItem = useMemo(() => {
    return items.find(i => i.id === selectedItemId) || items[0];
  }, [items, selectedItemId]);

  const toggleAuditFrequency = async (itemId: string, type: AuditType) => {
    const item = items.find(i => i.id === itemId);
    if (!item) return;

    const currentFreq = item.auditFrequency || [];
    const newFreq = currentFreq.includes(type)
      ? currentFreq.filter(f => f !== type)
      : [...currentFreq, type];

    try {
      await updateDoc(doc(db, 'inventory', itemId), { auditFrequency: newFreq });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `inventory/${itemId}`);
    }
  };

  const updateThreshold = async (itemId: string, threshold: number) => {
    try {
      await updateDoc(doc(db, 'inventory', itemId), { minThreshold: threshold });
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `inventory/${itemId}`);
    }
  };

  const handleDeleteAllData = async () => {
    if (deletePasswordInput === '112/58') {
      try {
        setIsSyncing(true);
        
        // Delete inventory
        const inventorySnapshot = await getDocs(collection(db, 'inventory'));
        let batch = writeBatch(db);
        let count = 0;
        for (const d of inventorySnapshot.docs) {
          batch.delete(d.ref);
          count++;
          if (count === 500) {
            await batch.commit();
            batch = writeBatch(db);
            count = 0;
          }
        }
        if (count > 0) await batch.commit();

        // Delete inbound logs
        const inboundSnapshot = await getDocs(collection(db, 'inbound_logs'));
        batch = writeBatch(db);
        count = 0;
        for (const d of inboundSnapshot.docs) {
          batch.delete(d.ref);
          count++;
          if (count === 500) {
            await batch.commit();
            batch = writeBatch(db);
            count = 0;
          }
        }
        if (count > 0) await batch.commit();

        // Delete audit logs
        const auditSnapshot = await getDocs(collection(db, 'audit_logs'));
        batch = writeBatch(db);
        count = 0;
        for (const d of auditSnapshot.docs) {
          batch.delete(d.ref);
          count++;
          if (count === 500) {
            await batch.commit();
            batch = writeBatch(db);
            count = 0;
          }
        }
        if (count > 0) await batch.commit();
        
        // Delete outbound logs
        const outboundSnapshot = await getDocs(collection(db, 'outbound_logs'));
        batch = writeBatch(db);
        count = 0;
        for (const d of outboundSnapshot.docs) {
          batch.delete(d.ref);
          count++;
          if (count === 500) {
            await batch.commit();
            batch = writeBatch(db);
            count = 0;
          }
        }
        if (count > 0) await batch.commit();
        
        // Delete backups
        const backupSnapshot = await getDocs(collection(db, 'backups'));
        batch = writeBatch(db);
        count = 0;
        for (const d of backupSnapshot.docs) {
          batch.delete(d.ref);
          count++;
          if (count === 500) {
            await batch.commit();
            batch = writeBatch(db);
            count = 0;
          }
        }
        if (count > 0) await batch.commit();
        
        // Reset settings to default
        await setDoc(doc(db, 'settings', 'global'), {
          globalThreshold: 10,
          customCategories: ['Thảo mộc thô', 'Thảo mộc khô', 'Pha chế', 'Vật tư', 'Vệ sinh'],
          customUnits: ['Gram', 'Kg', 'Cái', 'ml', 'Lưới', 'Thùng', 'Chai', 'Lon', 'Can'],
          gasWebhookUrl: DEFAULT_GAS_WEBHOOK
        }, { merge: true });

        setShowDeletePasswordModal(false);
        setDeletePasswordInput('');
        setPasswordError(false);
        
        setConfirmConfig({
          title: 'Đã xóa dữ liệu',
          message: 'Toàn bộ dữ liệu kho và nhật ký đã được xóa sạch.',
          onConfirm: () => setShowConfirmModal(false),
          icon: <Trash2 className="w-10 h-10 text-red-500" />,
          type: 'info'
        });
        setShowConfirmModal(true);
        navigateTo('overview');
      } catch (error) {
        handleFirestoreError(error, OperationType.DELETE, 'multiple collections');
      } finally {
        setIsSyncing(false);
      }
    } else {
      setPasswordError(true);
    }
  };

  const handleAddItem = async () => {
    if (!newItem.name) return;
    const id = Math.random().toString(36).substr(2, 9);
    const item: InventoryItem = {
      ...newItem as InventoryItem,
      id,
      status: 'Ổn định',
    };
    
    try {
      setIsSyncing(true);
      await setDoc(doc(db, 'inventory', id), item);
      
      await logAudit({
        id,
        name: item.name,
        type: 'Thêm mới',
        amount: item.actualStock,
        oldStock: 0,
        newStock: item.actualStock,
        reason: 'Thêm dược liệu mới vào hệ thống',
        timestamp: formatDate(new Date()),
        user: currentUser?.name || 'Unknown',
        details: `Thêm mới dược liệu ${item.name} với tồn kho ban đầu ${item.actualStock}`
      });

      // Background auto-sync to Sheets
      fetch('/api/sync/sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ inventory: [...items, item] }) // optimistically push new item
      }).catch(err => console.error('Background sync failed:', err));

      if (gasWebhookUrl && item.actualStock !== undefined) {
        fetch(gasWebhookUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain' },
          body: JSON.stringify({
            action: 'log_inventory_change',
            logs: [{
              user: currentUser?.name || 'Unknown',
              type: 'Thêm dược liệu mới',
              name: item.name,
              category: item.category,
              amount: item.actualStock,
              unit: item.unit,
              price: item.importPrice || 0,
              reason: 'Tồn kho ban đầu khai báo',
              stockAfter: item.actualStock
            }]
          })
        }).catch(err => console.error('Background GAS sync failed:', err));
      }

      navigateTo('inventory');
      setNewItem({
        name: '',
        category: 'Thảo mộc thô',
        unit: 'Gram',
        actualStock: 0,
        minThreshold: 0,
        maxThreshold: 0,
        description: '',
        location: '',
      });
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, 'inventory');
    } finally {
      setIsSyncing(false);
    }
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError('');
    setIsLoggingIn(true);

    try {
      const response = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: loginUsername, password: loginPassword })
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Tài khoản hoặc mật khẩu không chính xác');
      }
      
      const user = await response.json();

      if (user) {
        // Sign in to Firebase anonymously to satisfy security rules
        try {
          await signInAnonymously(auth);
        } catch (authErr: any) {
          if (authErr.code === 'auth/admin-restricted-operation') {
            console.error('LỖI QUAN TRỌNG: Bạn cần bật "Anonymous Authentication" trong Firebase Console.');
            throw new Error('Hệ thống chưa bật Anonymous Auth. Vui lòng bật trong Firebase Console để tiếp tục.');
          }
          throw authErr;
        }
        
        const userData: User = {
          uid: user.uid || user.username,
          username: user.username,
          role: user.role as Role,
          name: user.name || user.username
        };
        
        // Sync user profile to Firestore so rules can check role
        if (auth.currentUser) {
          await setDoc(doc(db, 'users', auth.currentUser.uid), {
            uid: auth.currentUser.uid,
            username: userData.username,
            role: userData.role,
            name: userData.name
          }, { merge: true });
        }

        setCurrentUser(userData);
        localStorage.setItem('botanical_user', JSON.stringify(userData));
        
        // Log login to GAS if configured
        if (gasWebhookUrl) {
          fetch(gasWebhookUrl, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              action: 'login',
              username: userData.username,
              role: userData.role,
              details: `Đăng nhập thành công lúc ${formatDate(new Date())}`
            })
          }).catch(err => console.error('GAS Login Log failed:', err));
        }

        const initialScreen = userData.role === 'quản lí' ? 'overview' : 'audit';
        setCurrentScreen(initialScreen);
        setScreenHistory([initialScreen]);
      }
    } catch (err: any) {
      console.error('Login Error:', err);
      setLoginError(err.message || 'Lỗi đăng nhập. Vui lòng thử lại.');
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    try {
      await auth.signOut();
      setCurrentUser(null);
      localStorage.removeItem('botanical_user');
      setCurrentScreen('login');
      setScreenHistory(['login']);
      setLoginUsername('');
      setLoginPassword('');
    } catch (error) {
      console.error('Logout Error:', error);
    }
  };

  useEffect(() => {
    if (currentScreen === 'details' && selectedItemId && firebaseUser) {
      const qInbound = query(collection(db, 'inbound_logs'), where('id', '==', selectedItemId));
      const qOutbound = query(collection(db, 'outbound_logs'), where('id', '==', selectedItemId));
      const qAudit = query(collection(db, 'audit_logs'), where('id', '==', selectedItemId));

      const unsubInbound = onSnapshot(qInbound, (snapshot) => {
        const inboundLogs = snapshot.docs.map(doc => ({ ...doc.data(), type: 'Nhập kho' } as LogEntry));
        setItemTransactions(prev => {
          const otherLogs = prev.filter(l => l.type !== 'Nhập kho');
          return [...otherLogs, ...inboundLogs].sort((a, b) => b.timestamp.localeCompare(a.timestamp));
        });
      }, (error) => handleFirestoreError(error, OperationType.GET, 'inbound_logs'));

      const unsubOutbound = onSnapshot(qOutbound, (snapshot) => {
        const outboundLogs = snapshot.docs.map(doc => ({ ...doc.data(), type: 'Xuất kho' } as LogEntry));
        setItemTransactions(prev => {
          const otherLogs = prev.filter(l => l.type !== 'Xuất kho');
          return [...otherLogs, ...outboundLogs].sort((a, b) => b.timestamp.localeCompare(a.timestamp));
        });
      }, (error) => handleFirestoreError(error, OperationType.GET, 'outbound_logs'));

      const unsubAudit = onSnapshot(qAudit, (snapshot) => {
        const auditLogs = snapshot.docs.map(doc => ({ ...doc.data(), type: 'Kiểm kê' } as LogEntry));
        setItemTransactions(prev => {
          const otherLogs = prev.filter(l => l.type !== 'Kiểm kê');
          return [...otherLogs, ...auditLogs].sort((a, b) => b.timestamp.localeCompare(a.timestamp));
        });
      }, (error) => handleFirestoreError(error, OperationType.GET, 'audit_logs'));

      return () => {
        unsubInbound();
        unsubOutbound();
        unsubAudit();
      };
    } else {
      setItemTransactions([]);
    }
  }, [currentScreen, selectedItemId, firebaseUser]);

  const chartData = useMemo(() => {
    if (!itemTransactions || itemTransactions.length === 0) return [];
    
    const parseLogDate = (dateStr: string) => {
      const match = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4}) - (\d{2}):(\d{2}):(\d{2})/);
      if (match) {
          return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]), Number(match[4]), Number(match[5]), Number(match[6])).getTime();
      }
      return new Date(dateStr).getTime() || 0;
    };

    const sortedLogs = [...itemTransactions].sort((a, b) => parseLogDate(a.timestamp) - parseLogDate(b.timestamp));
    
    const aggregated: Record<string, { date: string, stock: number, in: number, out: number }> = {};
    let currentStock = 0;
    
    sortedLogs.forEach(log => {
      const match = log.timestamp.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      const dayKey = match ? `${match[1]}/${match[2]}` : log.timestamp.split(' - ')[0];
      
      if (!aggregated[dayKey]) {
          aggregated[dayKey] = { date: dayKey, stock: currentStock, in: 0, out: 0 };
      }
      
      if (log.type === 'Nhập kho') {
          aggregated[dayKey].in += (log.amount || 0);
      } else if (log.type === 'Xuất kho') {
          aggregated[dayKey].out += (log.amount || 0);
      }
      
      if (log.newStock !== undefined) {
         currentStock = log.newStock;
      } else if (log.actualStock !== undefined) {
         currentStock = log.actualStock;
      }
      aggregated[dayKey].stock = currentStock;
    });

    return Object.values(aggregated);
  }, [itemTransactions]);

  const fifoItem = useMemo(() => {
    const itemsWithExpiry = items.filter(i => i.expiryDate && i.actualStock > 0);
    if (itemsWithExpiry.length === 0) return null;
    return itemsWithExpiry.sort((a, b) => (a.expiryDate || '').localeCompare(b.expiryDate || ''))[0];
  }, [items]);

  const filteredLogsForReport = useMemo(() => {
    return allLogs.filter(log => {
      if (!log.timestamp || typeof log.timestamp !== 'string') return false;
      const parts = log.timestamp.split(' - ');
      const datePart = parts[0];
      if (!datePart) return false;
      const dateParts = datePart.split('/');
      if (dateParts.length < 3) return false;
      // Ensure month and day are padded with 0
      const d = dateParts[0].padStart(2, '0');
      const m = dateParts[1].padStart(2, '0');
      const y = dateParts[2];
      const logDate = `${y}-${m}-${d}`;
      return logDate >= activeReportStartDate && logDate <= activeReportEndDate;
    });
  }, [allLogs, activeReportStartDate, activeReportEndDate]);

  const groupedLogsForReport = useMemo(() => {
    const groups: Record<string, typeof allLogs> = {};
    filteredLogsForReport.forEach(log => {
      const datePart = (log.timestamp || '').split(' - ')[0] || 'N/A';
      if (!groups[datePart]) {
        groups[datePart] = [];
      }
      groups[datePart].push(log);
    });
    
    const sortedDates = Object.keys(groups).sort((a, b) => {
      if (a === 'N/A') return 1;
      if (b === 'N/A') return -1;
      const partsA = a.split('/');
      const partsB = b.split('/');
      if (partsA.length === 3 && partsB.length === 3) {
        return new Date(`${partsB[2]}-${partsB[1]}-${partsB[0]}`).getTime() - new Date(`${partsA[2]}-${partsA[1]}-${partsA[0]}`).getTime();
      }
      return 0;
    });
    
    return sortedDates.map(date => ({
      date,
      logs: groups[date].sort((a, b) => {
        const timeA = (a.timestamp || '').split(' - ')[1] || '';
        const timeB = (b.timestamp || '').split(' - ')[1] || '';
        return timeB.localeCompare(timeA);
      })
    }));
  }, [filteredLogsForReport]);

  const filteredInboundLogs = useMemo(() => {
    return allLogs.filter(log => {
      if (log.type !== 'Nhập kho' || !log.timestamp || typeof log.timestamp !== 'string') return false;
      const parts = log.timestamp.split(' - ');
      const datePart = parts[0];
      if (!datePart) return false;
      const dateParts = datePart.split('/');
      if (dateParts.length < 3) return false;
      // Ensure month and day are padded with 0
      const d = dateParts[0].padStart(2, '0');
      const m = dateParts[1].padStart(2, '0');
      const y = dateParts[2];
      const logDate = `${y}-${m}-${d}`;
      return logDate >= inboundStartDate && logDate <= inboundEndDate;
    });
  }, [allLogs, inboundStartDate, inboundEndDate]);

  const filteredOutboundLogs = useMemo(() => {
    return allLogs.filter(log => {
      if (log.type !== 'Xuất kho' || !log.timestamp || typeof log.timestamp !== 'string') return false;
      const parts = log.timestamp.split(' - ');
      const datePart = parts[0];
      if (!datePart) return false;
      const dateParts = datePart.split('/');
      if (dateParts.length < 3) return false;
      // Ensure month and day are padded with 0
      const d = dateParts[0].padStart(2, '0');
      const m = dateParts[1].padStart(2, '0');
      const y = dateParts[2];
      const logDate = `${y}-${m}-${d}`;
      return logDate >= outboundStartDate && logDate <= outboundEndDate;
    });
  }, [allLogs, outboundStartDate, outboundEndDate]);

  const lowStockItems = useMemo(() => {
    return items.filter(item => item.minThreshold !== undefined && item.actualStock < item.minThreshold);
  }, [items]);

  if (!isAuthReady) {
    return (
      <div className="min-h-screen bg-surface flex items-center justify-center">
        <IOSSpinner />
      </div>
    );
  }

  return (
    <div className={`min-h-screen bg-slate-50 flex flex-col items-center transition-colors duration-500`}>
      <OfflineBanner isOnline={isOnline} />
      <LoadingOverlay isSyncing={isSyncing} />
      <Sidebar 
        isMenuOpen={isMenuOpen} 
        setIsMenuOpen={setIsMenuOpen} 
        currentUser={currentUser} 
        currentScreen={currentScreen} 
        navigateTo={navigateTo} 
        handleLogout={handleLogout} 
      />
      <ConfirmModal 
        isOpen={confirmModal.isOpen} 
        title={confirmModal.title} 
        message={confirmModal.message} 
        onConfirm={confirmModal.onConfirm} 
        onClose={() => setConfirmModal(prev => ({ ...prev, isOpen: false }))} 
      />

      {/* Edit Log Modal */}
      <AnimatePresence>
        {editingLog && (
          <>
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setEditingLog(null)}
              className="fixed inset-0 z-[100] bg-black/40 backdrop-blur-sm"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="fixed top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[90%] max-w-sm z-[110] bg-white dark:bg-[#1C1C1E] rounded-3xl shadow-2xl overflow-hidden border border-black/5"
            >
              <div className="p-6 space-y-6">
                <div className="text-center space-y-2">
                  <div className="w-12 h-12 bg-blue-500/10 text-blue-500 rounded-full flex items-center justify-center mx-auto mb-4">
                    <Edit3 size={24} />
                  </div>
                  <h3 className="text-xl font-bold text-[#1C1C1E] dark:text-white">Chỉnh sửa {editingLog.type.toLowerCase()}</h3>
                  <p className="text-sm text-apple-gray">{editingLog.name}</p>
                </div>

                <div className="space-y-4">
                  <div className="space-y-2">
                    <label className="text-xs font-bold uppercase tracking-widest text-apple-gray ml-2">Số lượng mới</label>
                    <input
                      type="number"
                      value={editLogAmount}
                      onChange={(e) => setEditLogAmount(e.target.value)}
                      className="w-full bg-black/5 dark:bg-white/5 border-none rounded-2xl p-4 text-[#1C1C1E] dark:text-white font-bold focus:ring-2 focus:ring-blue-500 transition-all"
                      placeholder="Nhập số lượng..."
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-bold uppercase tracking-widest text-apple-gray ml-2">Lý do chỉnh sửa</label>
                    <input
                      type="text"
                      value={editLogReason}
                      onChange={(e) => setEditLogReason(e.target.value)}
                      className="w-full bg-black/5 dark:bg-white/5 border-none rounded-2xl p-4 text-[#1C1C1E] dark:text-white font-bold focus:ring-2 focus:ring-blue-500 transition-all"
                      placeholder="Nhập lý do..."
                    />
                  </div>
                </div>

                <div className="flex gap-3 pt-2">
                  <button
                    onClick={() => setEditingLog(null)}
                    className="flex-1 py-3.5 rounded-xl font-bold text-[#1C1C1E] dark:text-white bg-black/5 dark:bg-white/5 active:scale-95 transition-all"
                  >
                    Hủy
                  </button>
                  <button
                    onClick={handleEditLogSubmit}
                    disabled={!editLogAmount || parseFloat(editLogAmount) <= 0}
                    className="flex-1 py-3.5 rounded-xl font-bold text-white bg-blue-500 active:scale-95 transition-all disabled:opacity-50 disabled:active:scale-100"
                  >
                    Lưu
                  </button>
                </div>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      <div className={`w-full max-w-md bg-[#F2F2F7] min-h-screen relative shadow-2xl md:shadow-none pb-12`}>
        
        {/* Top App Bar - Not fixed anymore */}
        {currentScreen !== 'login' && (
          <header className="w-full glass-nav px-6 py-8 relative overflow-hidden">
            {/* One-line herb decoration */}
            <div className="one-line-herb -top-4 -right-4 w-32 h-32 rotate-12">
              <svg viewBox="0 0 100 100" className="w-full h-full stroke-apple-green fill-none opacity-20">
                <path d="M50,90 Q50,50 80,20 M50,70 Q30,40 20,10 M50,80 Q70,50 90,40" strokeWidth="1" />
              </svg>
            </div>

            <div className="flex justify-between items-start w-full relative z-10">
              <div className="flex items-center gap-4">
                <button 
                  onClick={() => setIsMenuOpen(true)}
                  className="w-12 h-12 rounded-2xl bg-black/5 flex items-center justify-center active:scale-90 transition-all border border-black/5"
                >
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" className="text-[#1C1C1E]">
                    <line x1="3" y1="12" x2="21" y2="12"></line>
                    <line x1="3" y1="6" x2="21" y2="6"></line>
                    <line x1="3" y1="18" x2="21" y2="18"></line>
                  </svg>
                </button>
                <div className="flex flex-col gap-1">
                {currentScreen === 'overview' ? (
                  <>
                    <h1 className="text-[#1C1C1E] font-bold text-3xl tracking-tight">
                      {currentUser?.name || 'Nhân viên'}
                    </h1>
                    <div className="flex items-center gap-2 text-apple-gray text-sm font-medium">
                      <History className="w-3.5 h-3.5" />
                      <span>{formatDate(currentTime)}</span>
                    </div>
                  </>
                ) : (
                  <div className="flex items-center gap-4">
                    <button 
                      onClick={goBack}
                      className="w-10 h-10 rounded-full bg-black/5 flex items-center justify-center active:scale-90 transition-all"
                    >
                      <ArrowLeft className="w-5 h-5 text-[#1C1C1E]" />
                    </button>
                    <h1 className="text-[#1C1C1E] font-bold text-2xl tracking-tight">
                      {currentScreen === 'audit' && 'Kiểm kê'}
                      {currentScreen === 'details' && 'Chi tiết'}
                      {currentScreen === 'edit' && 'Chỉnh sửa'}
                      {currentScreen === 'report' && 'Báo cáo'}
                      {currentScreen === 'inventory' && 'Danh mục'}
                      {currentScreen === 'inbound' && 'Nhập kho'}
                      {currentScreen === 'settings' && 'Cài đặt'}
                      {currentScreen === 'add' && 'Thêm mới'}
                      {currentScreen === 'financial' && 'Tài chính'}
                    </h1>
                  </div>
                )}
              </div>
            </div>
            <div className="flex items-center gap-2">
                {currentUser?.role === 'quản lí' && currentScreen === 'overview' && (
                  <button 
                    onClick={() => navigateTo('settings')}
                    className="w-10 h-10 rounded-full bg-black/5 flex items-center justify-center active:scale-90 transition-all"
                  >
                    <Settings className="w-5 h-5 text-apple-gray" />
                  </button>
                )}
                <button 
                  onClick={handleLogout}
                  className="w-10 h-10 rounded-full bg-red-500/10 flex items-center justify-center active:scale-90 transition-all text-red-500"
                  title="Đăng xuất"
                >
                  <LogOut className="w-5 h-5" />
                </button>
              </div>
            </div>
          </header>
        )}

      <main className={`w-full ${currentScreen === 'login' ? 'p-0' : 'px-6 py-4'}`}>
        <PullToRefresh onRefresh={handleRefresh}>
          <AnimatePresence mode="wait">
          {currentScreen === 'login' && (
            <motion.div
              key="login"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="min-h-screen flex flex-col items-center justify-center bg-[#F2F2F7] p-8 relative overflow-hidden"
            >
              {/* One-line herb decoration */}
              <div className="one-line-herb top-10 left-10 w-64 h-64 opacity-10">
                <svg viewBox="0 0 100 100" className="w-full h-full stroke-apple-green fill-none">
                  <path d="M10,90 Q50,10 90,90 M30,70 Q50,30 70,70" strokeWidth="0.5" />
                </svg>
              </div>

              <div className="w-full max-w-sm space-y-12 relative z-10">
                <div className="text-center space-y-6">
                  <div className="w-24 h-24 bg-white rounded-[24px] flex items-center justify-center mx-auto shadow-xl border border-black/5">
                    <Leaf className="w-12 h-12 text-apple-green" />
                  </div>
                  <div>
                    <h1 className="text-3xl font-bold text-[#1C1C1E] tracking-tight border-0">Quản lí kho hàng</h1>
                    <p className="text-apple-gray font-medium mt-2">Hệ thống quản lý dược liệu</p>
                  </div>
                </div>

                <form onSubmit={handleLogin} className="space-y-8">
                  <div className="space-y-4">
                    <div className="apple-card p-1 bg-white/95 dark:bg-[#323234] shadow-2xl">
                      <div className="relative">
                        <div className="absolute inset-y-0 left-4 flex items-center pointer-events-none">
                          <UserIcon className="w-5 h-5 text-apple-gray" />
                        </div>
                        <input 
                          type="text"
                          required
                          value={loginUsername}
                          onChange={(e) => setLoginUsername(e.target.value)}
                          placeholder="Tên đăng nhập"
                          className="w-full bg-transparent border-none rounded-2xl py-4 pl-12 pr-6 text-[#1C1C1E] dark:text-white focus:ring-0 transition-all font-bold"
                        />
                      </div>
                      <div className="h-[1px] bg-black/5 dark:bg-white/10 mx-4"></div>
                      <div className="relative">
                        <div className="absolute inset-y-0 left-4 flex items-center pointer-events-none">
                          <Lock className="w-5 h-5 text-apple-gray" />
                        </div>
                        <input 
                          type="password"
                          required
                          value={loginPassword}
                          onChange={(e) => setLoginPassword(e.target.value)}
                          placeholder="Mật khẩu"
                          className="w-full bg-transparent border-none rounded-2xl py-4 pl-12 pr-6 text-[#1C1C1E] dark:text-white focus:ring-0 transition-all font-bold"
                        />
                      </div>
                    </div>
                  </div>

                  {loginError && (
                    <motion.div 
                      initial={{ opacity: 0, y: -10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="bg-red-500/10 border border-red-500/20 rounded-2xl p-4 flex items-center gap-3 text-red-500 text-sm font-semibold"
                    >
                      <AlertTriangle size={18} className="shrink-0" />
                      <span>{loginError}</span>
                    </motion.div>
                  )}

                  <button 
                    type="submit"
                    disabled={isLoggingIn}
                    className="w-full apple-button-primary flex items-center justify-center gap-2"
                  >
                    {isLoggingIn ? <IOSSpinner /> : (
                      <>
                        <span>Đăng nhập</span>
                        <ArrowRight className="w-5 h-5" />
                      </>
                    )}
                  </button>
                </form>

                <div className="text-center">
                  <p className="text-[10px] text-outline font-bold uppercase tracking-widest">
                    Botanical Archivist v2.0 • Secure Access
                  </p>
                </div>
              </div>
            </motion.div>
          )}

          {currentScreen === 'overview' && currentUser?.role === 'quản lí' && (
            <motion.div 
              key="overview"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-6"
            >
              {/* Low Stock Alerts */}
              {lowStockItems.length > 0 && (
                <div className="bg-red-50/80 border border-red-500/20 rounded-3xl p-6 relative overflow-hidden shadow-sm">
                  <div className="absolute -top-6 -right-6 p-8 opacity-[0.03]">
                    <AlertTriangle className="w-40 h-40 text-red-600" />
                  </div>
                  
                  <div className="flex flex-col gap-5 relative z-10">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-4">
                        <div className="w-12 h-12 rounded-2xl bg-red-500/10 flex items-center justify-center">
                          <AlertTriangle className="w-6 h-6 text-red-600" />
                        </div>
                        <div>
                          <h3 className="text-lg font-bold text-red-600">Cần nhập thêm hàng</h3>
                          <p className="text-sm font-medium text-red-600/80 mt-0.5">
                            {lowStockItems.length} mặt hàng dưới ngưỡng tồn kho an toàn
                          </p>
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-h-[300px] overflow-y-auto no-scrollbar pr-1 pb-1">
                      {lowStockItems.map(item => (
                        <div 
                          key={item.id}
                          onClick={() => {
                            setSelectedItemId(item.id);
                            navigateTo('details');
                          }}
                          className="bg-white rounded-2xl p-4 flex items-center gap-4 cursor-pointer shadow-sm border border-red-100/50 hover:border-red-500/30 hover:shadow-md transition-all group"
                        >
                          <div className="flex-1 min-w-0">
                            <h4 className="font-bold text-[#1C1C1E] truncate text-[15px]">{item.name}</h4>
                            <div className="flex items-center gap-2 mt-2">
                              <span className="text-xs font-bold text-red-600 bg-red-50 px-2.5 py-1 rounded-lg">
                                Tồn: {item.actualStock} {item.unit}
                              </span>
                              <div className="w-1 h-1 rounded-full bg-black/10"></div>
                              <span className="text-xs font-medium text-apple-gray flex items-center gap-1">
                                Ngưỡng: {item.minThreshold}
                              </span>
                            </div>
                          </div>
                          <div className="w-8 h-8 rounded-full bg-black/5 flex items-center justify-center group-hover:bg-red-50 transition-colors">
                            <ChevronRight className="w-4 h-4 text-apple-gray group-hover:text-red-500 transition-colors" />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {/* Hero Section */}
              <div className="grid grid-cols-1 gap-6">
                <div className="apple-card bg-gradient-to-br from-apple-blue to-[#5856D6] relative overflow-hidden flex min-h-[220px]">
                  {/* One-line herb decoration */}
                  <div className="one-line-herb -bottom-8 -right-8 w-48 h-48 rotate-45">
                    <svg viewBox="0 0 100 100" className="w-full h-full stroke-white fill-none opacity-20">
                      <path d="M10,90 Q40,40 90,10 M30,80 Q50,60 70,40 M50,90 Q70,70 90,50" strokeWidth="0.5" />
                    </svg>
                  </div>

                  <div className="flex-1 p-8 z-10 flex flex-col justify-between text-white">
                    <div>
                      <div className="flex items-center gap-2 mb-4">
                        <div className="w-8 h-8 bg-white/20 rounded-lg flex items-center justify-center">
                          <BarChart3 className="w-4 h-4 text-white" />
                        </div>
                        <h2 className="font-semibold text-xs uppercase tracking-wider opacity-80">Sức khỏe kho hàng</h2>
                      </div>
                      <div className="mt-2">
                        <div className="text-6xl font-bold tracking-tighter">{inventoryHealth}%</div>
                        <p className="text-white/70 mt-4 font-medium text-sm">Tổng tồn: {totalStockWeight} {items[0]?.unit || 'đơn vị'}</p>
                      </div>
                    </div>
                  </div>
                  
                  <div className="absolute right-8 bottom-8 z-20">
                    <button 
                      onClick={() => navigateTo('add')}
                      className="w-14 h-14 bg-white text-apple-blue rounded-2xl flex items-center justify-center shadow-xl active:scale-90 transition-all duration-300"
                    >
                      <Plus size={28} strokeWidth={3} />
                    </button>
                  </div>
                </div>
              </div>

              {/* Categories */}
              <section className="space-y-6">
                <div className="flex items-center justify-between px-2">
                  <div className="flex items-center gap-2">
                    <div className="w-1.5 h-5 bg-apple-green rounded-full"></div>
                    <h2 className="text-sm font-bold text-[#1C1C1E]">Danh mục chính</h2>
                  </div>
                  <button 
                    onClick={() => navigateTo('inventory')}
                    className="text-[11px] font-bold text-apple-blue bg-apple-blue/5 px-3 py-1 rounded-full active:scale-95 transition-all"
                  >
                    Xem tất cả
                  </button>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  {[
                    { name: 'Thảo mộc thô', icon: Leaf, count: items.filter(i => i.category === 'Thảo mộc thô').length },
                    { name: 'Pha chế', icon: Coffee, count: items.filter(i => i.category === 'Pha chế').length },
                    { name: 'Vật tư', icon: Boxes, count: items.filter(i => i.category === 'Vật tư').length },
                    { name: 'Vệ sinh', icon: Droplets, count: items.filter(i => i.category === 'Vệ sinh').length },
                  ].map((cat) => (
                    <motion.div 
                      key={cat.name} 
                      whileHover={{ y: -4 }}
                      whileTap={{ scale: 0.98 }}
                      onClick={() => setSearchQuery(cat.name)}
                      className={`relative overflow-hidden p-6 apple-card flex flex-col justify-between min-h-[140px] transition-all duration-300 border-2 ${searchQuery === cat.name ? 'border-apple-blue bg-apple-blue/5' : 'border-transparent'}`}
                    >
                      {/* One-line herb decoration */}
                      <div className="one-line-herb -top-2 -right-2 w-16 h-16 opacity-5">
                        <svg viewBox="0 0 100 100" className="w-full h-full stroke-apple-green fill-none">
                          <path d="M50,10 Q50,50 90,90 M10,50 Q50,50 90,10" strokeWidth="1" />
                        </svg>
                      </div>

                      <div className={`w-12 h-12 rounded-2xl flex items-center justify-center transition-colors ${searchQuery === cat.name ? 'bg-apple-blue text-white' : 'bg-apple-blue/5'}`}>
                        <cat.icon className={`w-6 h-6 ${searchQuery === cat.name ? 'text-white' : 'text-apple-blue'}`} />
                      </div>
                      <div className="relative z-10">
                        <span className={`block font-bold text-sm transition-colors ${searchQuery === cat.name ? 'text-apple-blue' : 'text-[#1C1C1E]'}`}>{cat.name}</span>
                        <span className="text-[11px] font-medium text-apple-gray">{cat.count} mặt hàng</span>
                      </div>
                    </motion.div>
                  ))}
                </div>
              </section>

              {/* Recent Items */}
              <section>
                <div className="apple-card overflow-hidden">
                  <div className="p-6 bg-white/40 border-b border-black/5 flex justify-between items-center">
                    <div className="flex items-center gap-3">
                      <Archive className="w-5 h-5 text-apple-gray" />
                      <h3 className="font-bold text-[#1C1C1E]">
                        {searchQuery ? `Kết quả cho "${searchQuery}"` : `Kho vật tư (${items.length} mục)`}
                      </h3>
                    </div>
                    {searchQuery && (
                      <button 
                        onClick={() => setSearchQuery('')}
                        className="text-[11px] font-bold text-apple-blue bg-apple-blue/10 px-3 py-1 rounded-full"
                      >
                        Xóa lọc
                      </button>
                    )}
                  </div>
                  <div className="divide-y divide-black/5">
                    {filteredItems.length > 0 ? (
                      filteredItems.map((item) => (
                        <div 
                          key={item.id} 
                          onClick={() => navigateTo('details', item.id)}
                          className="p-6 flex items-center justify-between hover:bg-black/5 transition-colors cursor-pointer"
                        >
                          <div className="flex items-center gap-4">
                            <div>
                              <h4 className="font-bold text-[#1C1C1E]">{item.name}</h4>
                              <p className="text-xs text-apple-gray truncate max-w-[150px]">{item.description}</p>
                            </div>
                          </div>
                          <div className="text-right">
                            <span className={`block font-bold text-lg ${item.status === 'Tồn thấp' ? 'text-red-500' : 'text-apple-blue'}`}>
                              {item.actualStock} {item.unit?.toLowerCase() || ''}
                            </span>
                            <span className={`text-[10px] font-bold uppercase tracking-tighter ${item.status === 'Tồn thấp' ? 'text-red-500' : 'text-apple-gray'}`}>
                              {item.status === 'Tồn thấp' ? 'Cần nhập gấp' : `Vị trí: ${item.location || 'N/A'}`}
                            </span>
                          </div>
                        </div>
                      ))
                    ) : (
                      <div className="p-12 text-center">
                        <div className="w-16 h-16 bg-black/5 rounded-full flex items-center justify-center mx-auto mb-4">
                          <Search className="w-8 h-8 text-apple-gray" />
                        </div>
                        <p className="text-apple-gray font-medium">Không tìm thấy dược liệu nào khớp với "{searchQuery}"</p>
                      </div>
                    )}
                  </div>
                </div>
              </section>
            </motion.div>
          )}

          {currentScreen === 'audit' && (
            <motion.div 
              key="audit"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-6 pb-24"
            >
              {/* Audit Header & Manage Toggle */}
              <section className="flex items-center justify-between px-2">
                <div>
                  <h2 className="text-2xl font-bold text-[#1C1C1E]">Kiểm kê {auditType === 'daily' ? 'Hàng ngày' : auditType === 'weekly' ? 'Hàng tuần' : 'Hàng tháng'}</h2>
                  <p className="text-apple-gray text-xs font-medium mt-1">Cập nhật số lượng tồn kho thực tế</p>
                </div>
                <button 
                  onClick={(e) => {
                    e.stopPropagation();
                    setIsEditingAudit(!isEditingAudit);
                  }}
                  className={`px-4 py-2 rounded-full text-[11px] font-bold transition-all flex items-center gap-2 z-20 ${
                    isEditingAudit 
                      ? 'bg-apple-blue text-white shadow-lg shadow-apple-blue/20' 
                      : 'bg-black/5 text-apple-gray'
                  }`}
                >
                  <Edit3 className="w-3.5 h-3.5" />
                  {isEditingAudit ? 'Xong' : 'Sửa'}
                </button>
              </section>

              {/* Audit Type Selector - Segmented Control Style */}
              <section className="bg-black/5 p-1 rounded-2xl flex gap-1">
                {(['daily', 'weekly', 'monthly'] as const).map((type) => (
                  <button
                    key={type}
                    onClick={() => {
                      setAuditType(type);
                      setAuditCategory('Tất cả');
                      setAuditSearchQuery('');
                    }}
                    className={`flex-1 py-2.5 rounded-xl text-[11px] font-bold transition-all ${
                      auditType === type 
                        ? 'bg-white text-[#1C1C1E] shadow-sm' 
                        : 'text-apple-gray hover:bg-black/5'
                    }`}
                  >
                    {type === 'daily' ? 'Ngày' : type === 'weekly' ? 'Tuần' : 'Tháng'}
                  </button>
                ))}
              </section>

              {/* Search Bar for Audit */}
              <div className="relative group">
                <div className="absolute inset-y-0 left-4 flex items-center pointer-events-none">
                  <Search className="w-4 h-4 text-apple-gray group-focus-within:text-apple-blue transition-colors" />
                </div>
                <input 
                  type="text"
                  placeholder="Tìm dược liệu..."
                  value={auditSearchQuery}
                  onChange={(e) => setAuditSearchQuery(e.target.value)}
                  className="w-full bg-black/5 border-none rounded-2xl py-3 pl-10 pr-10 text-[#1C1C1E] text-sm focus:ring-0 transition-all"
                />
                {auditSearchSuggestions.length > 0 && (
                  <div className="absolute top-full left-0 right-0 mt-2 bg-white/80 backdrop-blur-xl rounded-2xl shadow-2xl z-[60] overflow-hidden border border-black/5">
                    {auditSearchSuggestions.map(item => (
                      <button 
                        key={item.id}
                        onClick={() => {
                          setAuditSearchQuery('');
                          const element = document.getElementById(`audit-item-${item.id}`);
                          if (element) {
                            element.scrollIntoView({ behavior: 'smooth', block: 'center' });
                            element.classList.add('ring-2', 'ring-apple-blue');
                            setTimeout(() => element.classList.remove('ring-2', 'ring-apple-blue'), 2000);
                          }
                        }}
                        className="w-full px-6 py-4 flex items-center gap-4 hover:bg-apple-blue/5 transition-colors text-left border-b border-black/5 last:border-0"
                      >
                        <div className="flex-1">
                          <p className="text-[#1C1C1E] font-bold text-sm">{item.name}</p>
                          <p className="text-apple-gray text-[10px] font-bold uppercase tracking-widest">{item.category}</p>
                        </div>
                        <ArrowRight size={16} className="text-apple-gray" />
                      </button>
                    ))}
                  </div>
                )}
                {auditSearchQuery && (
                  <button 
                    onClick={() => setAuditSearchQuery('')}
                    className="absolute inset-y-0 right-4 flex items-center"
                  >
                    <X className="w-4 h-4 text-apple-gray hover:text-apple-blue" />
                  </button>
                )}
              </div>

              {/* Category Filter */}
              <div className="flex gap-2 overflow-x-auto no-scrollbar scroll-smooth pb-2">
                {['Tất cả', 'Thảo mộc thô', 'Pha chế', 'Vật tư', 'Vệ sinh'].map((cat) => (
                  <button
                    key={cat}
                    onClick={() => setAuditCategory(cat)}
                    className={`whitespace-nowrap px-4 py-2 rounded-full text-[11px] font-bold transition-all ${
                      auditCategory === cat
                        ? 'bg-apple-blue text-white shadow-lg shadow-apple-blue/20'
                        : 'bg-black/5 text-apple-gray'
                    }`}
                  >
                    {cat}
                  </button>
                ))}
              </div>

              {!isEditingAudit && (
                <section className="space-y-4">
                  <div className="apple-card p-6 relative overflow-hidden">
                    {/* One-line herb decoration */}
                    <div className="one-line-herb -top-2 -right-2 w-24 h-24 opacity-10">
                      <svg viewBox="0 0 100 100" className="w-full h-full stroke-apple-green fill-none">
                        <path d="M50,90 Q50,50 10,10 M50,70 Q70,40 90,10" strokeWidth="0.5" />
                      </svg>
                    </div>

                    <div className="flex items-end justify-between relative z-10">
                      <div>
                        <span className="text-[10px] font-bold text-apple-gray uppercase tracking-wider mb-1 block">Tiến độ hoàn thành</span>
                        <h2 className="text-5xl font-bold tracking-tighter text-apple-blue">
                          {Math.round((items.filter(i => i.auditFrequency?.includes(auditType) && (auditCategory === 'Tất cả' || i.category === auditCategory)).length / items.filter(i => i.auditFrequency?.includes(auditType)).length) * 100) || 0}%
                        </h2>
                      </div>
                      <div className="text-right">
                        <span className="text-xs font-bold text-[#1C1C1E]">
                          {items.filter(i => i.auditFrequency?.includes(auditType) && (auditCategory === 'Tất cả' || i.category === auditCategory)).length} mục tiêu
                        </span>
                        <p className="text-[10px] text-apple-gray font-medium">Cần hoàn thành hôm nay</p>
                      </div>
                    </div>
                    
                    <button 
                      onClick={handlePreAudit}
                      className="w-full apple-button-primary mt-6 flex items-center justify-center gap-2"
                    >
                      <span>Xác nhận Kiểm kê</span>
                      <ArrowRight className="w-4 h-4" />
                    </button>
                  </div>
                </section>
              )}

              <div className="grid grid-cols-1 gap-4">
                {filteredAuditItems.length > 0 ? (
                  filteredAuditItems.map((item, _index) => (
                    <div 
                      key={item.id} 
                      onClick={() => isEditingAudit && toggleAuditFrequency(item.id, auditType)}
                      className={`apple-card p-6 flex flex-col justify-between min-h-[160px] transition-all duration-300 relative ${
                        isEditingAudit 
                          ? item.auditFrequency?.includes(auditType)
                            ? 'border-apple-blue ring-2 ring-apple-blue shadow-lg shadow-apple-blue/10'
                            : 'opacity-60 grayscale-[0.5]'
                          : 'border-transparent'
                      } ${isEditingAudit ? 'cursor-pointer hover:scale-[1.02]' : ''}`}
                    >
                      <div className="flex justify-between items-start mb-4">
                        <div className="flex items-center gap-4">
                          <div>
                            <h3 className="text-lg font-bold text-[#1C1C1E]">{item.name}</h3>
                            <p className="text-xs text-apple-gray font-medium">{item.category} • {item.unit}</p>
                          </div>
                        </div>
                        {!isEditingAudit && (
                          <div className="text-right">
                            <span className="text-[10px] font-bold text-apple-gray uppercase block mb-1">Lý thuyết</span>
                            <span className="text-xl font-bold text-apple-gray/30">
                              {currentUser?.role === 'quản lí' ? item.actualStock : '***'}
                            </span>
                          </div>
                        )}
                        {isEditingAudit && (
                          <div className={`flex items-center justify-center w-10 h-10 rounded-2xl transition-all ${
                            item.auditFrequency?.includes(auditType) 
                              ? 'bg-apple-blue text-white shadow-md shadow-apple-blue/20' 
                              : 'bg-black/5 text-apple-gray'
                          }`}>
                            {item.auditFrequency?.includes(auditType) ? (
                              <CheckCircle2 className="w-6 h-6" />
                            ) : (
                              <Plus className="w-6 h-6" />
                            )}
                          </div>
                        )}
                      </div>
                      {isEditingAudit && (
                        <div className="mt-2">
                          <span className={`text-[10px] font-bold uppercase tracking-widest px-3 py-1 rounded-full ${
                            item.auditFrequency?.includes(auditType)
                              ? 'bg-apple-blue/10 text-apple-blue'
                              : 'bg-black/5 text-apple-gray'
                          }`}>
                            {item.auditFrequency?.includes(auditType) ? 'Đã thêm vào danh sách' : 'Chưa thêm'}
                          </span>
                        </div>
                      )}
                      {!isEditingAudit && (
                        <div className="relative flex items-center">
                          <input 
                            className="w-full bg-black/5 border-none rounded-2xl py-4 pl-6 pr-16 text-3xl font-bold text-apple-blue focus:ring-0 transition-all" 
                            type="number" 
                            placeholder="0"
                            value={auditInputs[item.id] === 0 ? '' : (auditInputs[item.id] || '')}
                            onChange={(e) => {
                              const val = e.target.value;
                              setAuditInputs(prev => ({
                                ...prev,
                                [item.id]: val === '' ? 0 : Number(val)
                              }));
                            }}
                          />
                          <span className="absolute right-6 text-xs font-bold text-apple-gray uppercase">{item.unit}</span>
                        </div>
                      )}
                    </div>
                  ))
                ) : (
                  <div className="p-12 text-center apple-card">
                    <div className="w-16 h-16 bg-black/5 rounded-full flex items-center justify-center mx-auto mb-4">
                      <Search className="w-8 h-8 text-apple-gray" />
                    </div>
                    <p className="text-apple-gray font-medium">Không tìm thấy dược liệu nào</p>
                  </div>
                )}
              </div>

              {/* Fixed bottom container removed from here */}
            </motion.div>
          )}

          {currentScreen === 'details' && selectedItem && (
            <motion.div 
              key="details"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 1.05 }}
              className="space-y-8 pb-24"
            >
              <section className="relative h-[240px] min-h-[240px] w-full rounded-[40px] overflow-hidden shadow-2xl bg-black/5 flex items-center px-8 text-apple-gray">
                <div className="absolute inset-0 bg-gradient-to-br from-apple-blue to-[#5856D6] flex flex-col justify-end p-8 pb-12">
                  <div className="flex items-center gap-2 mb-4">
                    <span className="bg-white/20 backdrop-blur-md text-white px-3 py-1.5 rounded-xl text-[10px] font-bold uppercase tracking-wider border border-white/20">{selectedItem.category}</span>
                    <span className="bg-white/20 backdrop-blur-md text-white px-3 py-1.5 rounded-xl text-[10px] font-bold uppercase tracking-wider border border-white/20">{selectedItem.unit}</span>
                  </div>
                  <h2 className="text-white text-4xl font-extrabold tracking-tight leading-loose">{selectedItem.name}</h2>
                </div>
                <div className="absolute top-0 right-0 w-48 h-48 -mr-12 -mt-12 opacity-10">
                   <svg viewBox="0 0 100 100" className="w-full h-full stroke-white fill-none">
                    <circle cx="50" cy="50" r="40" strokeWidth="0.5" />
                    <line x1="10" y1="50" x2="90" y2="50" strokeWidth="0.5" />
                  </svg>
                </div>
                <button 
                  onClick={() => navigateTo('edit', selectedItem.id)}
                  className="absolute top-6 right-6 p-3 bg-white/20 backdrop-blur-md rounded-full border border-white/20 text-white active:scale-90 transition-all"
                >
                  <Edit3 size={20} />
                </button>
              </section>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="md:col-span-1 apple-card p-6 flex flex-col justify-between">
                  <div>
                    <div className="flex items-center gap-2 mb-6">
                      <div className="w-8 h-8 rounded-full bg-apple-blue/10 flex items-center justify-center text-apple-blue">
                        <ClipboardList size={18} />
                      </div>
                      <h3 className="text-[#1C1C1E] font-bold text-lg">Quản lý FIFO</h3>
                    </div>
                    <div className="space-y-4">
                      <div className="bg-black/5 rounded-2xl p-4">
                        <p className="text-[10px] font-bold tracking-widest uppercase text-apple-gray mb-1">Số Lô Ưu Tiên</p>
                        <p className="text-xl font-bold text-apple-blue">{selectedItem.batchId || 'N/A'}</p>
                      </div>
                      <div className="bg-red-500/5 rounded-2xl p-4 border-l-4 border-red-500">
                        <p className="text-[10px] font-bold tracking-widest uppercase text-red-500/60 mb-1">Hạn Sử Dụng</p>
                        <p className="text-xl font-bold text-red-500">{selectedItem.expiryDate || 'N/A'}</p>
                      </div>
                    </div>
                  </div>
                  <p className="mt-6 text-xs text-apple-gray font-medium leading-relaxed">Ưu tiên xuất kho từ lô này để đảm bảo chất lượng dược tính tối ưu.</p>
                </div>

                <div className="md:col-span-2 apple-card p-8 relative overflow-hidden">
                  <div className="relative z-10">
                    <div className="flex items-center justify-between mb-8">
                      <div>
                        <h3 className="text-[#1C1C1E] font-bold text-2xl tracking-tight">Kiểm kho</h3>
                        <p className="text-apple-gray text-sm font-medium mt-1">Xác nhận khối lượng thực tế</p>
                      </div>
                      <div className="w-12 h-12 rounded-2xl bg-apple-blue/10 flex items-center justify-center text-apple-blue">
                        <BarChart3 size={24} />
                      </div>
                    </div>
                    <div className="space-y-6">
                      <div className="relative">
                        <label className="block text-[10px] font-bold tracking-widest uppercase text-apple-gray mb-3 ml-2">Nhập số lượng thực tế</label>
                        <div className="flex items-center bg-black/5 rounded-[32px] px-8 py-6 transition-all focus-within:bg-white focus-within:shadow-xl focus-within:ring-4 focus-within:ring-apple-blue/10">
                          <input 
                            className="bg-transparent border-none focus:ring-0 w-full text-4xl font-bold text-[#1C1C1E] placeholder:text-apple-gray/30" 
                            placeholder="0.00" 
                            type="number"
                            value={auditInputs[selectedItem.id] === 0 ? '' : (auditInputs[selectedItem.id] || '')}
                            onChange={(e) => {
                              const val = e.target.value;
                              setAuditInputs(prev => ({
                                ...prev,
                                [selectedItem.id]: val === '' ? 0 : Number(val)
                              }));
                            }}
                          />
                          <span className="text-2xl font-bold text-apple-gray">{selectedItem.unit}</span>
                        </div>
                      </div>
                      {currentUser?.role === 'quản lí' && (
                        <div className="bg-apple-blue/5 rounded-2xl p-4 flex items-center gap-4">
                          <div className="w-6 h-6 rounded-full bg-apple-blue/20 flex items-center justify-center">
                            <Info size={14} className="text-apple-blue" />
                          </div>
                          <p className="text-xs text-apple-blue font-medium italic">Chế độ kiểm kho mù yêu cầu nhập giá trị thực tế mà không hiển thị tồn kho hệ thống.</p>
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                {currentUser?.role === 'quản lí' && (
                  <div className="md:col-span-3 apple-card p-8">
                    <div className="flex items-center justify-between mb-8">
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 flex items-center justify-center text-apple-blue">
                          <Settings size={20} />
                        </div>
                        <h3 className="text-[#1C1C1E] font-bold text-xl">Cài đặt cảnh báo</h3>
                      </div>
                    </div>
                    <div className="bg-black/5 rounded-3xl p-6 border border-apple-blue/10">
                      <label className="block text-[10px] font-bold tracking-widest uppercase text-apple-gray mb-4 ml-2">Ngưỡng tồn kho tối thiểu</label>
                      <div className="flex items-center gap-4">
                        <div className="flex-1 flex items-center bg-white rounded-2xl px-6 py-4 shadow-sm border border-black/5 focus-within:ring-4 focus-within:ring-apple-blue/10 transition-all">
                          <input 
                            className="bg-transparent border-none focus:ring-0 w-full text-2xl font-bold text-[#1C1C1E] placeholder:text-apple-gray/30" 
                            placeholder="0" 
                            type="number"
                            value={selectedItem.minThreshold === 0 ? '' : (selectedItem.minThreshold || '')}
                            onChange={(e) => {
                              const val = e.target.value === '' ? 0 : Number(e.target.value);
                              updateThreshold(selectedItem.id, val);
                            }}
                          />
                          <span className="text-lg font-bold text-apple-gray">{selectedItem.unit}</span>
                        </div>
                      </div>
                      <p className="mt-4 text-xs text-apple-gray font-medium italic">Hệ thống sẽ hiển thị cảnh báo nổi bật khi số lượng thực tế thấp hơn giá trị này.</p>
                    </div>
                  </div>
                )}

                {chartData.length > 0 && (
                  <div className="md:col-span-3 apple-card p-8">
                    <div className="flex items-center gap-3 mb-8">
                      <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 flex items-center justify-center text-apple-blue">
                        <TrendingUp size={20} />
                      </div>
                      <h3 className="text-[#1C1C1E] font-bold text-xl">Biểu đồ biến động kho</h3>
                    </div>
                    <div className="w-full h-80">
                      <ResponsiveContainer width="100%" height="100%">
                        <LineChart data={chartData} margin={{ top: 5, right: 20, bottom: 5, left: 0 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" vertical={false} />
                          <XAxis dataKey="date" stroke="#9ca3af" fontSize={12} tickLine={false} axisLine={false} />
                          <YAxis stroke="#9ca3af" fontSize={12} tickLine={false} axisLine={false} />
                          <RechartsTooltip 
                            contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.1)', fontWeight: 500 }}
                            itemStyle={{ fontWeight: 700 }}
                          />
                          <Legend wrapperStyle={{ paddingTop: '20px' }} iconType="circle" />
                          <Line type="monotone" name="Tồn kho" dataKey="stock" stroke="#007AFF" strokeWidth={3} dot={{ r: 4, strokeWidth: 2 }} activeDot={{ r: 6 }} />
                          <Line type="monotone" name="Nhập kho" dataKey="in" stroke="#34C759" strokeWidth={2} dot={false} strokeDasharray="4 4" />
                          <Line type="monotone" name="Xuất kho" dataKey="out" stroke="#FF3B30" strokeWidth={2} dot={false} strokeDasharray="4 4" />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}

                <div className="md:col-span-3 apple-card p-8">
                  <div className="flex items-center justify-between mb-8">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 flex items-center justify-center text-apple-blue">
                        <History size={20} />
                      </div>
                      <h3 className="text-[#1C1C1E] font-bold text-xl">Lịch sử giao dịch</h3>
                    </div>
                    <button 
                      onClick={() => navigateTo('inventory')}
                      className="text-apple-blue font-bold text-sm hover:underline active:scale-95 transition-all"
                    >
                      Xem tất cả
                    </button>
                  </div>
                  <div className="space-y-3">
                    {itemTransactions.length > 0 ? (
                      itemTransactions.slice(0, 10).map((t, idx) => (
                        <div key={t.id + idx} className="bg-black/5 rounded-2xl p-5 flex items-center justify-between transition-all hover:bg-black/10">
                          <div className="flex items-center gap-4">
                            <div className={`w-12 h-12 rounded-full flex items-center justify-center ${t.type === 'Kiểm kê' ? 'bg-apple-blue/10 text-apple-blue' : 'bg-apple-green/10 text-apple-green'}`}>
                              {t.type === 'Kiểm kê' ? <ClipboardList className="w-6 h-6" /> : <PlusCircleIcon className="w-6 h-6" />}
                            </div>
                            <div>
                              <p className="font-bold text-[#1C1C1E]">{t.type} {t.reason ? `(${t.reason})` : ''}</p>
                              <p className="text-[10px] text-apple-gray font-bold uppercase tracking-wider">{t.timestamp}</p>
                              <p className="text-[10px] text-apple-gray font-medium mt-0.5">Thực hiện bởi: {t.user}</p>
                            </div>
                          </div>
                          <div className="text-right">
                            <p className={`text-xl font-bold ${t.type === 'Kiểm kê' ? 'text-apple-blue' : 'text-apple-green'}`}>
                              {t.type === 'Kiểm kê' ? '' : '+'}{t.amount || ((t.actualStock ?? 0) - (t.theoreticalStock ?? 0))} {selectedItem.unit?.toLowerCase() || ''}
                            </p>
                            {t.type === 'Kiểm kê' && (
                              <p className="text-[10px] text-apple-gray font-bold uppercase tracking-wider">Lệch: {(t.actualStock ?? 0) - (t.theoreticalStock ?? 0)}</p>
                            )}
                          </div>
                        </div>
                      ))
                    ) : (
                      <div className="p-8 text-center text-apple-gray font-medium italic">Chưa có giao dịch nào được ghi nhận.</div>
                    )}
                  </div>
                </div>
              </div>

              <div className="flex justify-center pt-4">
                <button 
                  onClick={() => setShowConfirmModal(true)}
                  className="apple-button-primary w-full max-w-md py-5 flex items-center justify-center gap-3"
                >
                  <Save className="w-5 h-5" />
                  Lưu kiểm kê
                </button>
              </div>
            </motion.div>
          )}

          {currentScreen === 'inbound' && (
            <motion.div 
              key="inbound"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-8 pb-24"
            >
              <section className="px-2">
                <h2 className="text-3xl font-bold text-[#1C1C1E] tracking-tight">Phiếu Nhập Kho</h2>
                <div className="flex bg-black/5 rounded-xl p-1 mt-4 w-fit">
                  <button 
                    onClick={() => setInboundTab('record')}
                    className={`px-6 py-2 rounded-lg text-sm font-bold transition-all ${inboundTab === 'record' ? 'bg-white text-[#1C1C1E] shadow-sm' : 'text-apple-gray'}`}
                  >
                    Ghi nhận
                  </button>
                  <button 
                    onClick={() => setInboundTab('history')}
                    className={`px-6 py-2 rounded-lg text-sm font-bold transition-all ${inboundTab === 'history' ? 'bg-white text-[#1C1C1E] shadow-sm' : 'text-apple-gray'}`}
                  >
                    Lịch sử
                  </button>
                </div>
              </section>

              {inboundTab === 'record' && (
              <div className="space-y-6">
                <div className="apple-card p-6 space-y-6">
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Chọn nguyên liệu</label>
                    <div className="relative group">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-apple-gray group-focus-within:text-apple-blue transition-colors" />
                      <input 
                        type="text"
                        placeholder="Tìm dược liệu nhập kho..."
                        value={inboundSearchQuery}
                        onChange={(e) => setInboundSearchQuery(e.target.value)}
                        className="w-full bg-black/5 border-none rounded-2xl h-14 pl-12 pr-4 text-[#1C1C1E] font-bold focus:ring-0 transition-all"
                      />
                      {inboundSearchSuggestions.length > 0 && (
                        <div className="absolute top-full left-0 right-0 mt-2 bg-white/80 backdrop-blur-xl rounded-2xl shadow-2xl z-[60] overflow-hidden border border-black/5">
                          {inboundSearchSuggestions.map(item => (
                            <button 
                              key={item.id}
                              onClick={() => {
                                setInboundSearchQuery('');
                                setCurrentInbound(prev => ({ ...prev, itemId: item.id }));
                              }}
                              className="w-full px-6 py-4 flex items-center gap-4 hover:bg-apple-blue/5 transition-colors text-left border-b border-black/5 last:border-0"
                            >
                              <div className="flex-1">
                                <p className="text-[#1C1C1E] font-bold text-sm">{item.name}</p>
                                <p className="text-apple-gray text-[10px] font-bold uppercase tracking-widest">{item.category}</p>
                              </div>
                              <CheckCircle2 size={16} className="text-apple-blue" />
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                    {currentInbound.itemId && (
                      <div className="mt-2 p-3 bg-apple-blue/5 rounded-xl flex items-center justify-between border border-apple-blue/10">
                        <div className="flex items-center gap-3">
                          <span className="text-sm font-bold text-apple-blue">{items.find(i => i.id === currentInbound.itemId)?.name}</span>
                        </div>
                        <button onClick={() => setCurrentInbound(prev => ({ ...prev, itemId: '' }))} className="text-apple-gray hover:text-red-500">
                          <X size={16} />
                        </button>
                      </div>
                    )}
                  </div>

                    <div className="space-y-2">
                      <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Số lượng nhập</label>
                      <div className="relative">
                        <input 
                          type="number"
                          placeholder="0.00"
                          className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-2xl font-bold text-apple-blue focus:ring-0"
                          value={currentInbound.amount === 0 ? '' : (currentInbound.amount || '')}
                          onChange={(e) => {
                            const val = e.target.value;
                            setCurrentInbound(prev => ({ ...prev, amount: val === '' ? 0 : Number(val) }));
                          }}
                          min="0.01"
                          step="0.01"
                        />
                        <span className="absolute right-4 top-1/2 -translate-y-1/2 font-bold text-apple-gray uppercase text-xs">
                          {items.find(i => i.id === currentInbound.itemId)?.unit || ''}
                        </span>
                      </div>
                    </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Ghi chú</label>
                    <textarea 
                      className="w-full bg-black/5 border-none rounded-2xl p-4 text-[#1C1C1E] font-medium focus:ring-0 min-h-[80px]"
                      placeholder="Ví dụ: Nhập hàng từ nhà cung cấp A..."
                      value={currentInbound.reason}
                      onChange={(e) => setCurrentInbound(prev => ({ ...prev, reason: e.target.value }))}
                    />
                  </div>

                  <button 
                    type="button"
                    onClick={addToInboundCart}
                    disabled={!currentInbound.itemId || currentInbound.amount <= 0}
                    className="w-full apple-button-secondary py-4 flex items-center justify-center gap-2"
                  >
                    <Plus className="w-5 h-5" />
                    Thêm vào danh sách
                  </button>
                </div>

                {inboundCart.length > 0 && (
                  <div className="space-y-4">
                    <div className="flex items-center justify-between px-2">
                      <h3 className="text-[10px] font-bold uppercase tracking-widest text-apple-gray">Danh sách chờ nhập ({inboundCart.length})</h3>
                      <button 
                        onClick={() => setInboundCart([])}
                        className="text-[10px] font-bold text-red-500 uppercase tracking-widest hover:underline"
                      >
                        Xóa tất cả
                      </button>
                    </div>
                    <div className="space-y-2">
                      {inboundCart.map((cartItem, index) => {
                        const item = items.find(i => i.id === cartItem.itemId);
                        return (
                          <div key={index} className="apple-card p-4 flex items-center justify-between">
                            <div className="flex items-center gap-4">
                              <div className="w-10 h-10 rounded-xl bg-apple-blue/10 flex items-center justify-center text-apple-blue font-bold">
                                {item?.name.charAt(0)}
                              </div>
                              <div>
                                <p className="font-bold text-[#1C1C1E]">{item?.name}</p>
                                <p className="text-[10px] text-apple-gray font-bold">+{cartItem.amount} {item?.unit} • {cartItem.reason}</p>
                              </div>
                            </div>
                            <button 
                              onClick={() => removeFromInboundCart(index)}
                              className="p-2 text-red-500 hover:bg-red-500/10 rounded-xl transition-colors"
                            >
                              <Trash2 className="w-5 h-5" />
                            </button>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                <div className="flex gap-4">
                  <button 
                    type="button"
                    onClick={() => navigateTo('inventory')}
                    className="flex-1 apple-button-secondary"
                  >
                    Hủy bỏ
                  </button>
                  <button 
                    type="button"
                    onClick={handleInboundSubmit}
                    disabled={isSyncing || inboundCart.length === 0}
                    className="flex-[2] apple-button-primary flex items-center justify-center gap-2"
                  >
                    {isSyncing ? <IOSSpinner /> : (
                      <>
                        <CheckCircle2 className="w-5 h-5" />
                        Xác nhận nhập ({inboundCart.length})
                      </>
                    )}
                  </button>
                </div>
              </div>
              )}

              {inboundTab === 'history' && (
                <div className="apple-card p-6 space-y-6">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-2xl bg-apple-green/10 text-apple-green flex items-center justify-center">
                        <History size={20} />
                      </div>
                      <h3 className="text-[#1C1C1E] font-bold text-xl">Lịch sử nhập kho</h3>
                    </div>
                    <div className="flex items-center gap-2 bg-black/5 p-1 rounded-xl">
                      <input 
                        type="date" 
                        value={inboundStartDate}
                        onChange={(e) => setInboundStartDate(e.target.value)}
                        className="bg-transparent border-none text-[10px] font-bold text-[#1C1C1E] focus:ring-0 p-1 w-24"
                      />
                      <span className="text-apple-gray text-[10px] font-bold">→</span>
                      <input 
                        type="date" 
                        value={inboundEndDate}
                        onChange={(e) => setInboundEndDate(e.target.value)}
                        className="bg-transparent border-none text-[10px] font-bold text-[#1C1C1E] focus:ring-0 p-1 w-24"
                      />
                    </div>
                  </div>
                  
                  <div className="space-y-3 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                    {filteredInboundLogs.length > 0 ? (
                      filteredInboundLogs.map((log, i) => (
                        <SwipeableLogItem
                          key={log.docId || i}
                          onDelete={() => handleDeleteLog(log)}
                          onEdit={() => {
                            setEditingLog(log);
                            setEditLogAmount(log.amount?.toString() || '');
                            setEditLogReason(log.reason || '');
                          }}
                          disabled={currentUser?.role !== 'quản lí'}
                        >
                          <div className="bg-black/5 rounded-2xl p-4 flex items-center justify-between">
                            <div className="flex items-center gap-4">
                              <div className="w-10 h-10 rounded-xl bg-apple-green/10 text-apple-green flex items-center justify-center">
                                <Plus size={18} />
                              </div>
                              <div>
                                <h4 className="font-bold text-[#1C1C1E] text-sm">{items.find(item => item.id === log.id)?.name || log.name || 'Không xác định'}</h4>
                                <p className="text-[10px] text-apple-gray font-medium">{log.user} • {log.timestamp}</p>
                              </div>
                            </div>
                            <div className="text-right">
                              <p className="text-sm font-bold text-apple-green">+{log.amount}</p>
                              <p className="text-[10px] text-apple-gray font-bold uppercase tracking-widest">{log.unit}</p>
                            </div>
                          </div>
                        </SwipeableLogItem>
                      ))
                    ) : (
                      <p className="text-center text-apple-gray text-xs py-8 italic">Không có dữ liệu trong khoảng thời gian này</p>
                    )}
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {currentScreen === 'outbound' && (
            <motion.div 
              key="outbound"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-8 pb-24"
            >
              <section className="px-2">
                <h2 className="text-3xl font-bold text-[#1C1C1E] tracking-tight">Phiếu Xuất Kho</h2>
                <div className="flex bg-black/5 rounded-xl p-1 mt-4 w-fit">
                  <button 
                    onClick={() => setOutboundTab('record')}
                    className={`px-6 py-2 rounded-lg text-sm font-bold transition-all ${outboundTab === 'record' ? 'bg-white text-[#1C1C1E] shadow-sm' : 'text-apple-gray'}`}
                  >
                    Ghi nhận
                  </button>
                  <button 
                    onClick={() => setOutboundTab('history')}
                    className={`px-6 py-2 rounded-lg text-sm font-bold transition-all ${outboundTab === 'history' ? 'bg-white text-[#1C1C1E] shadow-sm' : 'text-apple-gray'}`}
                  >
                    Lịch sử
                  </button>
                </div>
              </section>

              {outboundTab === 'record' && (
              <div className="space-y-6">
                <div className="apple-card p-6 space-y-6">
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Chọn nguyên liệu</label>
                    <div className="relative group">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-apple-gray group-focus-within:text-apple-blue transition-colors" />
                      <input 
                        type="text"
                        placeholder="Tìm dược liệu xuất kho..."
                        value={outboundSearchQuery}
                        onChange={(e) => setOutboundSearchQuery(e.target.value)}
                        className="w-full bg-black/5 border-none rounded-2xl h-14 pl-12 pr-4 text-[#1C1C1E] font-bold focus:ring-0 transition-all"
                      />
                      {outboundSearchSuggestions.length > 0 && (
                        <div className="absolute top-full left-0 right-0 mt-2 bg-white/80 backdrop-blur-xl rounded-2xl shadow-2xl z-[60] overflow-hidden border border-black/5">
                          {outboundSearchSuggestions.map(item => (
                            <button 
                              key={item.id}
                              onClick={() => {
                                setOutboundSearchQuery('');
                                setCurrentOutbound(prev => ({ ...prev, itemId: item.id }));
                              }}
                              className="w-full px-6 py-4 flex items-center gap-4 hover:bg-apple-blue/5 transition-colors text-left border-b border-black/5 last:border-0"
                            >
                              <div className="flex-1">
                                <p className="text-[#1C1C1E] font-bold text-sm">{item.name}</p>
                                <p className="text-apple-gray text-[10px] font-bold uppercase tracking-widest">{item.category}</p>
                              </div>
                              <CheckCircle2 size={16} className="text-apple-blue" />
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                    {currentOutbound.itemId && (
                      <div className="mt-2 p-3 bg-apple-blue/5 rounded-xl flex items-center justify-between border border-apple-blue/10">
                        <div className="flex items-center gap-3">
                          <span className="text-sm font-bold text-apple-blue">{items.find(i => i.id === currentOutbound.itemId)?.name}</span>
                        </div>
                        <button onClick={() => setCurrentOutbound(prev => ({ ...prev, itemId: '' }))} className="text-apple-gray hover:text-red-500">
                          <X size={16} />
                        </button>
                      </div>
                    )}
                  </div>

                    <div className="space-y-2">
                      <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Số lượng xuất</label>
                      <div className="relative">
                        <input 
                          type="number"
                          placeholder="0.00"
                          className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-2xl font-bold text-apple-blue focus:ring-0"
                          value={currentOutbound.amount === 0 ? '' : (currentOutbound.amount || '')}
                          onChange={(e) => {
                            const val = e.target.value;
                            setCurrentOutbound(prev => ({ ...prev, amount: val === '' ? 0 : Number(val) }));
                          }}
                          min="0.01"
                          step="0.01"
                        />
                        <span className="absolute right-4 top-1/2 -translate-y-1/2 font-bold text-apple-gray uppercase text-xs">
                          {items.find(i => i.id === currentOutbound.itemId)?.unit || ''}
                        </span>
                      </div>
                    </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Lý do xuất</label>
                    <textarea 
                      className="w-full bg-black/5 border-none rounded-2xl p-4 text-[#1C1C1E] font-medium focus:ring-0 min-h-[80px]"
                      placeholder="Ví dụ: Xuất hàng bán lẻ, Xuất kho pha chế..."
                      value={currentOutbound.reason}
                      onChange={(e) => setCurrentOutbound(prev => ({ ...prev, reason: e.target.value }))}
                    />
                  </div>

                  <button 
                    type="button"
                    onClick={addToOutboundCart}
                    disabled={!currentOutbound.itemId || currentOutbound.amount <= 0}
                    className="w-full apple-button-secondary py-4 flex items-center justify-center gap-2"
                  >
                    <Plus className="w-5 h-5" />
                    Thêm vào danh sách
                  </button>
                </div>

                {outboundCart.length > 0 && (
                  <div className="space-y-4">
                    <div className="flex items-center justify-between px-2">
                      <h3 className="text-[10px] font-bold uppercase tracking-widest text-apple-gray">Danh sách chờ xuất ({outboundCart.length})</h3>
                      <button 
                        onClick={() => setOutboundCart([])}
                        className="text-[10px] font-bold text-red-500 uppercase tracking-widest hover:underline"
                      >
                        Xóa tất cả
                      </button>
                    </div>
                    <div className="space-y-2">
                      {outboundCart.map((cartItem, index) => {
                        const item = items.find(i => i.id === cartItem.itemId);
                        return (
                          <div key={index} className="apple-card p-4 flex items-center justify-between">
                            <div className="flex items-center gap-4">
                              <div className="w-10 h-10 rounded-xl bg-apple-blue/10 flex items-center justify-center text-apple-blue font-bold">
                                {item?.name.charAt(0)}
                              </div>
                              <div>
                                <p className="font-bold text-[#1C1C1E]">{item?.name}</p>
                                <p className="text-[10px] text-apple-gray font-bold">-{cartItem.amount} {item?.unit} • {cartItem.reason}</p>
                              </div>
                            </div>
                            <button 
                              onClick={() => removeFromOutboundCart(index)}
                              className="p-2 text-red-500 hover:bg-red-500/10 rounded-xl transition-colors"
                            >
                              <Trash2 className="w-5 h-5" />
                            </button>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                <div className="flex gap-4">
                  <button 
                    type="button"
                    onClick={() => navigateTo('inventory')}
                    className="flex-1 apple-button-secondary"
                  >
                    Hủy bỏ
                  </button>
                  <button 
                    type="button"
                    onClick={handleOutboundSubmit}
                    disabled={isSyncing || outboundCart.length === 0}
                    className="flex-[2] apple-button-primary flex items-center justify-center gap-2"
                  >
                    {isSyncing ? <IOSSpinner /> : (
                      <>
                        <CheckCircle2 className="w-5 h-5" />
                        Xác nhận xuất ({outboundCart.length})
                      </>
                    )}
                  </button>
                </div>
              </div>
              )}

              {outboundTab === 'history' && (
                <div className="apple-card p-6 space-y-6">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                        <History size={20} />
                      </div>
                      <h3 className="text-[#1C1C1E] font-bold text-xl">Lịch sử xuất kho</h3>
                    </div>
                    <div className="flex items-center gap-2 bg-black/5 p-1 rounded-xl">
                      <input 
                        type="date" 
                        value={outboundStartDate}
                        onChange={(e) => setOutboundStartDate(e.target.value)}
                        className="bg-transparent border-none text-[10px] font-bold text-[#1C1C1E] focus:ring-0 p-1 w-24"
                      />
                      <span className="text-apple-gray text-[10px] font-bold">→</span>
                      <input 
                        type="date" 
                        value={outboundEndDate}
                        onChange={(e) => setOutboundEndDate(e.target.value)}
                        className="bg-transparent border-none text-[10px] font-bold text-[#1C1C1E] focus:ring-0 p-1 w-24"
                      />
                    </div>
                  </div>
                  
                  <div className="space-y-3 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                    {filteredOutboundLogs.length > 0 ? (
                      filteredOutboundLogs.map((log, i) => (
                        <SwipeableLogItem
                          key={log.docId || i}
                          onDelete={() => handleDeleteLog(log)}
                          onEdit={() => {
                            setEditingLog(log);
                            setEditLogAmount(log.amount?.toString() || '');
                            setEditLogReason(log.reason || '');
                          }}
                          disabled={currentUser?.role !== 'quản lí'}
                        >
                          <div className="bg-black/5 rounded-2xl p-4 flex items-center justify-between">
                            <div className="flex items-center gap-4">
                              <div className="w-10 h-10 rounded-xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                                <Minus size={18} />
                              </div>
                              <div>
                                <h4 className="font-bold text-[#1C1C1E] text-sm">{items.find(item => item.id === log.id)?.name || log.name || 'Không xác định'}</h4>
                                <p className="text-[10px] text-apple-gray font-medium">{log.user} • {log.timestamp}</p>
                              </div>
                            </div>
                            <div className="text-right">
                              <p className="text-sm font-bold text-apple-blue">-{log.amount}</p>
                              <p className="text-[10px] text-apple-gray font-bold uppercase tracking-widest">{log.unit}</p>
                            </div>
                          </div>
                        </SwipeableLogItem>
                      ))
                    ) : (
                      <p className="text-center text-apple-gray text-xs py-8 italic">Không có dữ liệu trong khoảng thời gian này</p>
                    )}
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {currentScreen === 'edit' && selectedItem && editingItem && (
            <motion.div 
              key="edit"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8 pb-24"
            >
              <section className="px-2">
                <h2 className="text-3xl font-bold text-[#1C1C1E] tracking-tight">Chỉnh sửa</h2>
                <p className="text-apple-gray font-medium mt-1">{selectedItem.name}</p>
              </section>

              <div className="grid grid-cols-1 gap-6">
                <div className="apple-card p-6 space-y-6">
                    <div className="space-y-2">
                      <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Tên dược liệu</label>
                      <input 
                        className="w-full bg-black/5 border-none rounded-2xl h-14 px-5 text-[#1C1C1E] text-lg font-bold focus:ring-0 transition-all" 
                        type="text" 
                        value={editingItem.name || ''}
                        onChange={(e) => setEditingItem(prev => ({ ...prev, name: e.target.value }))}
                      />
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Nhóm danh mục</label>
                        <div className="relative">
                          <select 
                            className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0 appearance-none"
                            value={editingItem.category || ''}
                            onChange={(e) => setEditingItem(prev => ({ ...prev, category: e.target.value as any }))}
                          >
                            {customCategories.map(cat => (
                              <option key={cat} value={cat}>{cat}</option>
                            ))}
                          </select>
                          <ChevronDown className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-apple-gray" />
                        </div>
                      </div>

                      <div className="space-y-2">
                        <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Đơn vị tính</label>
                        <div className="flex flex-wrap gap-2">
                          {customUnits.map((u) => (
                            <button 
                              key={u}
                              type="button"
                              onClick={() => setEditingItem(prev => ({ ...prev, unit: u }))}
                              className={`px-4 py-3 rounded-xl text-xs font-bold transition-all active:scale-95 ${editingItem.unit === u ? 'bg-apple-blue text-white shadow-lg shadow-apple-blue/20' : 'bg-black/5 text-apple-gray hover:bg-black/10'}`}
                            >
                              {u}
                            </button>
                          ))}
                        </div>
                      </div>
                    </div>
                </div>

                <div className="apple-card p-6 space-y-6">
                  <h3 className="text-lg font-bold text-[#1C1C1E] px-2">Thông số lưu kho</h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Ngưỡng tối thiểu</label>
                        <input 
                          className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0" 
                          type="number" 
                          placeholder="0"
                          value={editingItem.minThreshold === 0 ? '' : (editingItem.minThreshold || '')}
                          onChange={(e) => {
                            const val = e.target.value;
                            setEditingItem(prev => ({ ...prev, minThreshold: val === '' ? 0 : Number(val) }));
                          }}
                        />
                      </div>
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Ngưỡng tối đa</label>
                        <input 
                          className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0" 
                          type="number" 
                          placeholder="0"
                          value={editingItem.maxThreshold === 0 ? '' : (editingItem.maxThreshold || '')}
                          onChange={(e) => {
                            const val = e.target.value;
                            setEditingItem(prev => ({ ...prev, maxThreshold: val === '' ? 0 : Number(val) }));
                          }}
                        />
                      </div>
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Giá nhập (VND / {editingItem.unit})</label>
                        <input 
                          className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0" 
                          type="number" 
                          placeholder="0"
                          value={editingItem.importPrice === 0 ? '' : (editingItem.importPrice || '')}
                          onChange={(e) => {
                            const val = e.target.value;
                            setEditingItem(prev => ({ ...prev, importPrice: val === '' ? 0 : Number(val) }));
                          }}
                        />
                      </div>
                    <div className="space-y-2 md:col-span-2">
                      <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Vị trí kệ</label>
                      <input 
                        className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0" 
                        type="text" 
                        value={editingItem.location || ''}
                        onChange={(e) => setEditingItem(prev => ({ ...prev, location: e.target.value }))}
                      />
                    </div>
                  </div>
                </div>
              </div>

              <div className="flex gap-4 pt-4">
                <button onClick={() => navigateTo('details')} className="flex-1 apple-button-secondary">Hủy bỏ</button>
                <button 
                  onClick={() => {
                    setConfirmConfig({
                      title: 'Xóa dược liệu',
                      message: `Bạn có chắc chắn muốn xóa "${selectedItem.name}" khỏi hệ thống không? Thao tác này không thể hoàn tác.`,
                      onConfirm: () => handleDeleteItem(selectedItem.id),
                      icon: <Trash2 className="w-10 h-10" />,
                      type: 'danger'
                    });
                    setShowConfirmModal(true);
                  }}
                  className="flex-1 py-4 bg-red-500/10 text-red-500 font-bold rounded-2xl flex items-center justify-center gap-2 hover:bg-red-500/20 transition-all"
                >
                  <Trash2 size={18} />
                  Xóa
                </button>
                <button 
                  onClick={async () => {
                    if (!editingItem) return;

                    const updatedData = {
                      name: editingItem.name,
                      category: editingItem.category,
                      minThreshold: editingItem.minThreshold,
                      maxThreshold: editingItem.maxThreshold,
                      importPrice: editingItem.importPrice,
                      location: editingItem.location,
                      unit: editingItem.unit
                    };

                    try {
                      setIsSyncing(true);
                      await updateDoc(doc(db, 'inventory', selectedItem.id), updatedData);
                      
                      // Log the change
                      await logAudit({
                        id: selectedItem.id,
                        name: selectedItem.name,
                        type: 'Cập nhật thông tin',
                        amount: 0,
                        oldStock: selectedItem.actualStock,
                        newStock: selectedItem.actualStock,
                        reason: 'Chỉnh sửa thông tin dược liệu',
                        timestamp: formatDate(new Date()),
                        user: currentUser?.name || 'Unknown',
                        details: `Cập nhật: ${Object.keys(updatedData).join(', ')}`
                      });

                      // Background auto-sync to Sheets
                      fetch('/api/sync/sheets', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ 
                          inventory: items.map(i => i.id === selectedItem.id ? { ...i, ...updatedData } : i) 
                        })
                      }).catch(err => console.error('Background sync failed:', err));

                      // Check if threshold increase causes low stock alert
                      const threshold = updatedData.minThreshold || globalThreshold;
                      if (selectedItem.actualStock > 0 && selectedItem.actualStock <= threshold && selectedItem.actualStock > (selectedItem.minThreshold || globalThreshold)) {
                        notifyManagers(
                          'Cảnh báo Tồn kho',
                          `Mặt hàng ${selectedItem.name} đang ở mức dưới cảnh báo do cập nhật ngưỡng tối thiểu.`,
                          '/suppliers'
                        );
                      }

                      navigateTo('details');
                    } catch (error) {
                      handleFirestoreError(error, OperationType.UPDATE, `inventory/${selectedItem.id}`);
                    } finally {
                      setIsSyncing(false);
                      setEditingItem(null);
                    }
                  }} 
                  disabled={isSyncing}
                  className="flex-[2] apple-button-primary flex items-center justify-center gap-2"
                >
                  {isSyncing ? <IOSSpinner /> : (
                    <>
                      <CheckCircle2 className="w-5 h-5" />
                      Lưu thay đổi
                    </>
                  )}
                </button>
              </div>
            </motion.div>
          )}

          {currentScreen === 'settings' && (
            <motion.div 
              key="settings"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-8 pb-24"
            >
              <section className="space-y-6">
                <div className="px-2">
                  <span className="text-[10px] font-bold tracking-widest uppercase text-apple-gray mb-1 block">Tùy chỉnh ứng dụng</span>
                  <h2 className="text-3xl font-bold text-[#1C1C1E]">Cài đặt hệ thống</h2>
                </div>

                {/* Push Notifications (Manager Only) */}
                {currentUser?.role === 'quản lí' && pushNotifications.isSupported && (
                  <div className="apple-card p-6 space-y-4 shadow-xl shadow-apple-blue/5">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <div className={`w-10 h-10 rounded-2xl flex items-center justify-center ${pushNotifications.isSubscribed ? 'bg-apple-blue/10 text-apple-blue' : 'bg-black/5 text-apple-gray'}`}>
                          <Bell size={20} />
                        </div>
                        <div>
                          <h4 className="font-bold text-[#1C1C1E]">Thông báo thiết bị</h4>
                          <p className="text-xs text-apple-gray font-medium">
                            {pushNotifications.isSubscribed ? 'Đang nhận thông báo' : 'Chưa bật thông báo'}
                          </p>
                        </div>
                      </div>
                      <button 
                        onClick={async () => {
                          if (pushNotifications.isSubscribed) {
                            await pushNotifications.unsubscribeFromPush();
                          } else {
                            await pushNotifications.subscribeToPush();
                          }
                        }}
                        className={`w-14 h-7 rounded-full relative transition-all duration-300 ${pushNotifications.isSubscribed ? 'bg-apple-blue' : 'bg-black/10'}`}
                      >
                        <div className={`absolute top-1 w-5 h-5 bg-white rounded-full shadow-sm transition-all duration-300 ${pushNotifications.isSubscribed ? 'left-8' : 'left-1'}`} />
                      </button>
                    </div>
                    {pushNotifications.permission === 'denied' && (
                      <p className="text-xs font-bold text-red-500 mt-2">
                        Bạn đã sửa cài đặt trình duyệt để từ chối thông báo. Vui lòng bật lại trong cài đặt.
                      </p>
                    )}
                  </div>
                )}

                {/* Global Threshold */}
                <div className="apple-card p-6 space-y-6">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-2xl bg-red-500/10 text-red-500 flex items-center justify-center">
                      <AlertTriangle size={20} />
                    </div>
                    <div>
                      <h4 className="font-bold text-[#1C1C1E]">Cảnh báo tồn kho chung</h4>
                      <p className="text-xs text-apple-gray font-medium">Ngưỡng báo động mặc định: {globalThreshold}%</p>
                    </div>
                  </div>
                  <div className="px-2">
                    <input 
                      type="range" 
                      min="1" 
                      max="50" 
                      value={globalThreshold}
                      onChange={async (e) => {
                        const val = parseInt(e.target.value);
                        setGlobalThreshold(val);
                        try {
                          await setDoc(doc(db, 'settings', 'global'), { globalThreshold: val }, { merge: true });
                        } catch (error) {
                          handleFirestoreError(error, OperationType.UPDATE, 'settings/global');
                        }
                      }}
                      className="w-full accent-apple-blue h-1.5 bg-black/5 rounded-full appearance-none cursor-pointer"
                    />
                    <div className="flex justify-between mt-2 text-[10px] font-bold text-apple-gray uppercase tracking-widest">
                      <span>1%</span>
                      <span>25%</span>
                      <span>50%</span>
                    </div>
                  </div>
                </div>

                {/* Categories & Units */}
                <div className="apple-card p-6 space-y-8">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                      <LayoutGrid size={20} />
                    </div>
                    <h4 className="font-bold text-[#1C1C1E]">Danh mục & Đơn vị</h4>
                  </div>
                  
                  <div className="space-y-6">
                    <div>
                      <p className="text-[10px] font-bold uppercase tracking-widest text-apple-gray mb-3 ml-2">Google Apps Script Webhook</p>
                      <div className="space-y-4">
                        <input 
                          type="text"
                          placeholder={DEFAULT_GAS_WEBHOOK}
                          className="w-full bg-black/5 border-none rounded-2xl h-12 px-4 text-xs font-medium focus:ring-0"
                          value={gasWebhookUrl}
                          onChange={async (e) => {
                            const url = e.target.value;
                            setGasWebhookUrl(url);
                            try {
                              await setDoc(doc(db, 'settings', 'global'), { gasWebhookUrl: url }, { merge: true });
                            } catch (error) {
                              handleFirestoreError(error, OperationType.UPDATE, 'settings/global');
                            }
                          }}
                        />
                        <p className="text-[10px] text-apple-gray italic px-2">
                          * Sử dụng Webhook để đẩy dữ liệu trực tiếp lên Google Sheets mà không cần Service Account.
                        </p>
                      </div>
                    </div>

                    <div>
                      <p className="text-[10px] font-bold uppercase tracking-widest text-apple-gray mb-3 ml-2">Danh mục ({customCategories.length})</p>
                      <div className="flex flex-wrap gap-2">
                        {customCategories.map(cat => (
                          <span key={cat} className="px-4 py-2 bg-black/5 rounded-full text-xs font-bold text-[#1C1C1E] flex items-center gap-2 border border-black/5">
                            {cat}
                            <X size={14} className="cursor-pointer text-apple-gray hover:text-red-500 transition-colors" onClick={async () => {
                              const newCats = customCategories.filter(c => c !== cat);
                              setCustomCategories(newCats);
                              try {
                                await setDoc(doc(db, 'settings', 'global'), { customCategories: newCats }, { merge: true });
                              } catch (error) {
                                handleFirestoreError(error, OperationType.UPDATE, 'settings/global');
                              }
                            }} />
                          </span>
                        ))}
                        {isAddingCategory ? (
                          <div className="flex items-center gap-2 w-full">
                            <input 
                              type="text"
                              autoFocus
                              placeholder="Tên danh mục..."
                              className="flex-1 bg-black/5 border-none rounded-full h-9 px-4 text-xs font-bold focus:ring-1 focus:ring-apple-blue"
                              value={newCategoryName}
                              onChange={(e) => setNewCategoryName(e.target.value)}
                              onKeyDown={async (e) => {
                                if (e.key === 'Enter') {
                                  if (newCategoryName.trim()) {
                                    const newCats = [...customCategories, newCategoryName.trim()];
                                    setCustomCategories(newCats);
                                    try {
                                      await setDoc(doc(db, 'settings', 'global'), { customCategories: newCats }, { merge: true });
                                    } catch (error) {
                                      handleFirestoreError(error, OperationType.UPDATE, 'settings/global');
                                    }
                                  }
                                  setIsAddingCategory(false);
                                  setNewCategoryName('');
                                } else if (e.key === 'Escape') {
                                  setIsAddingCategory(false);
                                  setNewCategoryName('');
                                }
                              }}
                            />
                            <button 
                              onClick={async () => {
                                if (newCategoryName.trim()) {
                                  const newCats = [...customCategories, newCategoryName.trim()];
                                  setCustomCategories(newCats);
                                  try {
                                    await setDoc(doc(db, 'settings', 'global'), { customCategories: newCats }, { merge: true });
                                  } catch (error) {
                                    handleFirestoreError(error, OperationType.UPDATE, 'settings/global');
                                  }
                                }
                                setIsAddingCategory(false);
                                setNewCategoryName('');
                              }}
                              className="p-2 bg-apple-blue text-white rounded-full active:scale-90 transition-all"
                            >
                              <Check size={14} />
                            </button>
                            <button 
                              onClick={() => {
                                setIsAddingCategory(false);
                                setNewCategoryName('');
                              }}
                              className="p-2 bg-black/5 text-apple-gray rounded-full active:scale-90 transition-all"
                            >
                              <X size={14} />
                            </button>
                          </div>
                        ) : (
                          <button 
                            onClick={() => setIsAddingCategory(true)}
                            className="px-4 py-2 border-2 border-dashed border-black/10 rounded-full text-xs font-bold text-apple-blue flex items-center gap-1 hover:bg-apple-blue/5 transition-all"
                          >
                            <Plus size={14} /> Thêm mới
                          </button>
                        )}
                      </div>
                    </div>
                    <div>
                      <p className="text-[10px] font-bold uppercase tracking-widest text-apple-gray mb-3 ml-2">Đơn vị ({customUnits.length})</p>
                      <div className="flex flex-wrap gap-2">
                        {customUnits.map(unit => (
                          <span key={unit} className="px-4 py-2 bg-black/5 rounded-full text-xs font-bold text-[#1C1C1E] flex items-center gap-2 border border-black/5">
                            {unit}
                            <X size={14} className="cursor-pointer text-apple-gray hover:text-red-500 transition-colors" onClick={async () => {
                              const newUnits = customUnits.filter(u => u !== unit);
                              setCustomUnits(newUnits);
                              try {
                                await setDoc(doc(db, 'settings', 'global'), { customUnits: newUnits }, { merge: true });
                              } catch (error) {
                                handleFirestoreError(error, OperationType.UPDATE, 'settings/global');
                              }
                            }} />
                          </span>
                        ))}
                        {isAddingUnit ? (
                          <div className="flex items-center gap-2 w-full">
                            <input 
                              type="text"
                              autoFocus
                              placeholder="Tên đơn vị..."
                              className="flex-1 bg-black/5 border-none rounded-full h-9 px-4 text-xs font-bold focus:ring-1 focus:ring-apple-blue"
                              value={newUnitName}
                              onChange={(e) => setNewUnitName(e.target.value)}
                              onKeyDown={async (e) => {
                                if (e.key === 'Enter') {
                                  if (newUnitName.trim()) {
                                    const newUnits = [...customUnits, newUnitName.trim()];
                                    setCustomUnits(newUnits);
                                    try {
                                      await setDoc(doc(db, 'settings', 'global'), { customUnits: newUnits }, { merge: true });
                                    } catch (error) {
                                      handleFirestoreError(error, OperationType.UPDATE, 'settings/global');
                                    }
                                  }
                                  setIsAddingUnit(false);
                                  setNewUnitName('');
                                } else if (e.key === 'Escape') {
                                  setIsAddingUnit(false);
                                  setNewUnitName('');
                                }
                              }}
                            />
                            <button 
                              onClick={async () => {
                                if (newUnitName.trim()) {
                                  const newUnits = [...customUnits, newUnitName.trim()];
                                  setCustomUnits(newUnits);
                                  try {
                                    await setDoc(doc(db, 'settings', 'global'), { customUnits: newUnits }, { merge: true });
                                  } catch (error) {
                                    handleFirestoreError(error, OperationType.UPDATE, 'settings/global');
                                  }
                                }
                                setIsAddingUnit(false);
                                setNewUnitName('');
                              }}
                              className="p-2 bg-apple-blue text-white rounded-full active:scale-90 transition-all"
                            >
                              <Check size={14} />
                            </button>
                            <button 
                              onClick={() => {
                                setIsAddingUnit(false);
                                setNewUnitName('');
                              }}
                              className="p-2 bg-black/5 text-apple-gray rounded-full active:scale-90 transition-all"
                            >
                              <X size={14} />
                            </button>
                          </div>
                        ) : (
                          <button 
                            onClick={() => setIsAddingUnit(true)}
                            className="px-4 py-2 border-2 border-dashed border-black/10 rounded-full text-xs font-bold text-apple-blue flex items-center gap-1 hover:bg-apple-blue/5 transition-all"
                          >
                            <Plus size={14} /> Thêm mới
                          </button>
                        )}
                      </div>
                    </div>
                  </div>
                </div>

                {/* Audit Log */}
                <div className="apple-card p-6 space-y-6">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                        <History size={20} />
                      </div>
                      <h4 className="font-bold text-[#1C1C1E]">Nhật ký hệ thống</h4>
                    </div>
                  </div>
                  <div className="space-y-3 max-h-[300px] overflow-y-auto pr-2 custom-scrollbar smooth-scroll">
                    {allLogs.slice(0, 10).map((log, i) => (
                      <div key={i} className="flex justify-between items-center py-3 border-b border-black/5 last:border-0">
                        <div>
                          <p className="text-sm font-bold text-[#1C1C1E]">{log.type}: {items.find(item => item.id === log.id)?.name || log.name || 'Không xác định'}</p>
                          <p className="text-[10px] text-apple-gray font-medium">{log.user} • {log.timestamp}</p>
                        </div>
                        <span className={`text-[10px] font-bold uppercase tracking-wider ${log.type === 'Nhập kho' ? 'text-apple-green' : 'text-apple-blue'}`}>
                          {log.amount ? `+${log.amount}` : 'Kiểm kê'}
                        </span>
                      </div>
                    ))}
                    {allLogs.length === 0 && (
                      <p className="text-center text-apple-gray text-xs py-4">Chưa có nhật ký hoạt động</p>
                    )}
                  </div>
                </div>

                {/* Export & Backup */}
                <div className="grid grid-cols-2 gap-4">
                  <button 
                    onClick={exportToExcel}
                    className="apple-card p-6 flex flex-col items-center gap-3 active:scale-95 transition-all hover:bg-black/5"
                  >
                    <div className="w-12 h-12 bg-apple-green/10 text-apple-green rounded-2xl flex items-center justify-center">
                      <FileText size={24} />
                    </div>
                    <span className="text-[10px] font-bold text-[#1C1C1E] uppercase tracking-widest text-center">Xuất Excel</span>
                  </button>
                  <button 
                    onClick={syncAndBackupData}
                    disabled={isSyncing}
                    className="apple-card p-6 flex flex-col items-center gap-3 active:scale-95 transition-all hover:bg-black/5 disabled:opacity-50"
                  >
                    <div className="w-12 h-12 bg-apple-blue/10 text-apple-blue rounded-2xl flex items-center justify-center">
                      {isSyncing ? <IOSSpinner /> : <Cloud size={24} />}
                    </div>
                    <span className="text-[10px] font-bold text-[#1C1C1E] uppercase tracking-widest text-center text-balance leading-tight">Đồng bộ &<br/>Sao lưu</span>
                  </button>
                </div>

                {currentUser?.role === 'quản lí' && (
                  <button 
                    onClick={() => {
                      setShowDeletePasswordModal(true);
                      setPasswordError(false);
                      setDeletePasswordInput('');
                    }}
                    className="w-full py-5 bg-red-500/10 text-red-500 font-bold rounded-[24px] flex items-center justify-center gap-2 active:scale-95 transition-all border border-red-500/20"
                  >
                    <Trash2 size={18} />
                    XÓA TOÀN BỘ DỮ LIỆU
                  </button>
                )}
              </section>
            </motion.div>
          )}

          {currentScreen === 'report' && (
            <motion.div 
              key="report"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8 pb-24"
            >
              <section className="px-2 flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div className="flex-1">
                  <h2 className="text-2xl font-bold text-[#1C1C1E]">Báo cáo kho</h2>
                  <p className="text-apple-gray text-xs font-medium mt-1">Phân tích biến động tồn kho</p>
                </div>
                <div className="flex-shrink-0 flex flex-wrap items-center gap-2 bg-black/5 p-2 rounded-2xl border border-black/5 relative z-30">
                  <div className="flex items-center gap-1 flex-wrap">
                    <input 
                      type="date" 
                      value={reportStartDate}
                      onChange={(e) => setReportStartDate(e.target.value)}
                      className="bg-transparent border-none text-xs font-bold text-[#1C1C1E] focus:ring-0 p-1 cursor-pointer min-w-[130px] relative z-40"
                    />
                    <span className="text-apple-gray text-[10px] font-bold">→</span>
                    <input 
                      type="date" 
                      value={reportEndDate}
                      onChange={(e) => setReportEndDate(e.target.value)}
                      className="bg-transparent border-none text-xs font-bold text-[#1C1C1E] focus:ring-0 p-1 cursor-pointer min-w-[130px] relative z-40"
                    />
                  </div>
                  <button 
                    onClick={() => {
                      setActiveReportStartDate(reportStartDate);
                      setActiveReportEndDate(reportEndDate);
                    }}
                    className="bg-apple-blue text-white px-4 py-2 rounded-xl text-xs font-bold shadow-lg shadow-apple-blue/20 active:scale-95 transition-all flex items-center gap-2"
                  >
                    <Search size={14} />
                    Truy vấn
                  </button>
                </div>
              </section>

              <section className="bg-red-500 rounded-[40px] p-8 relative overflow-hidden shadow-2xl shadow-red-500/20">
                {/* Decoration */}
                <div className="absolute -top-10 -right-10 w-40 h-40 bg-white/10 rounded-full blur-3xl" />
                
                <div className="flex justify-between items-start mb-8 relative z-10">
                  <div>
                    <span className="text-[10px] font-bold tracking-widest uppercase text-white/60 mb-2 block">Chỉ số chênh lệch</span>
                    <h2 className="text-6xl font-bold leading-none tracking-tighter text-white">-2.3<span className="text-2xl text-white/60 ml-1">%</span></h2>
                  </div>
                  <div className="w-16 h-16 bg-white/20 backdrop-blur-md rounded-[24px] flex items-center justify-center border border-white/20">
                    <AlertTriangle className="w-8 h-8 text-white" />
                  </div>
                </div>
                <div className="bg-white/20 backdrop-blur-md text-white text-[10px] font-bold uppercase tracking-widest px-4 py-2.5 rounded-full inline-flex items-center gap-2 border border-white/20">
                  <div className="w-2 h-2 rounded-full bg-white animate-pulse" />
                  Vượt ngưỡng an toàn 0.5%
                </div>
              </section>

              <div className="grid grid-cols-2 gap-4">
                <div className="apple-card p-6">
                  <p className="text-[10px] font-bold text-apple-gray uppercase tracking-widest mb-2">Tồn lý thuyết</p>
                  <p className="text-3xl font-bold text-[#1C1C1E]">12,450</p>
                  <p className="text-[10px] text-apple-gray font-medium mt-1">Đơn vị lưu kho</p>
                </div>
                <div className="apple-card p-6">
                  <p className="text-[10px] font-bold text-apple-gray uppercase tracking-widest mb-2">Tồn thực tế</p>
                  <p className="text-3xl font-bold text-[#1C1C1E]">12,164</p>
                  <p className="text-[10px] text-red-500 font-bold mt-1">-286 đơn vị hụt</p>
                </div>
              </div>

              <section className="apple-card p-8 relative overflow-hidden">
                {/* Herb pattern */}
                <div className="herb-pattern opacity-5" />
                
                <div className="flex items-center justify-between mb-8 relative z-10">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                      <Archive size={20} />
                    </div>
                    <h3 className="text-[#1C1C1E] font-bold text-xl">Gợi ý FIFO</h3>
                  </div>
                  <span className="bg-apple-blue/10 text-apple-blue text-[10px] px-3 py-1 rounded-full font-bold uppercase tracking-wider">Ưu tiên</span>
                </div>
                
                {fifoItem ? (
                  <>
                    <div className="bg-black/5 rounded-3xl p-6 flex items-center gap-6 mb-8 relative z-10">
                      <div>
                        <h4 className="font-bold text-[#1C1C1E] text-lg">{fifoItem.name}</h4>
                        <div className="flex flex-col gap-1 mt-1">
                          <p className="text-[10px] text-apple-gray font-bold uppercase tracking-wider">Tồn kho: {fifoItem.actualStock} {fifoItem.unit}</p>
                          <p className="text-[10px] text-red-500 font-bold uppercase tracking-wider">HSD: {fifoItem.expiryDate || 'N/A'}</p>
                        </div>
                      </div>
                    </div>
                    
                    <button 
                      onClick={() => {
                        setSelectedItemId(fifoItem.id);
                        navigateTo('outbound');
                      }}
                      className="w-full apple-button-primary py-5 flex items-center justify-center gap-2 relative z-10"
                    >
                      <PlusCircle className="w-5 h-5" />
                      XUẤT KHO ƯU TIÊN
                    </button>
                  </>
                ) : (
                  <div className="text-center py-8">
                    <Archive size={40} className="mx-auto text-apple-gray/20 mb-4" />
                    <p className="text-apple-gray text-sm font-medium">Không có gợi ý FIFO nào</p>
                  </div>
                )}
              </section>

              <section className="apple-card p-8">
                <div className="flex items-center justify-between mb-8">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                      <History size={20} />
                    </div>
                    <h3 className="text-[#1C1C1E] font-bold text-xl">Nhật ký tồn kho</h3>
                  </div>
                  <span className="text-[10px] font-bold text-apple-gray uppercase tracking-widest">{filteredLogsForReport.length} bản ghi</span>
                </div>
                
                <div className="space-y-8 max-h-[600px] overflow-y-auto pr-2 custom-scrollbar">
                  {groupedLogsForReport.map((group, idx) => (
                    <div key={idx} className="bg-white rounded-[24px] border border-black/5 overflow-hidden shadow-sm">
                      <div 
                        className="bg-black/5 px-6 py-4 border-b border-black/5 flex items-center justify-between cursor-pointer hover:bg-black/10 transition-colors"
                        onClick={() => setCollapsedReportGroups(prev => ({ ...prev, [group.date]: !prev[group.date] }))}
                      >
                        <div className="flex items-center gap-2">
                          <div className="w-8 h-8 rounded-full bg-white flex items-center justify-center text-apple-gray shadow-sm">
                            <Clock size={14} />
                          </div>
                          <h4 className="font-bold text-[#1C1C1E]">{group.date}</h4>
                        </div>
                        <div className="flex items-center gap-3">
                          <span className="text-[10px] font-bold uppercase tracking-widest text-apple-gray bg-white px-3 py-1 rounded-full shadow-sm">
                            {group.logs.length} bản ghi
                          </span>
                          <ChevronDown size={18} className={`text-apple-gray transition-transform duration-200 ${collapsedReportGroups[group.date] ? 'rotate-180' : ''}`} />
                        </div>
                      </div>
                      <AnimatePresence initial={false}>
                        {!collapsedReportGroups[group.date] && (
                          <motion.div
                            initial={{ height: 0, opacity: 0 }}
                            animate={{ height: 'auto', opacity: 1 }}
                            exit={{ height: 0, opacity: 0 }}
                            transition={{ duration: 0.2 }}
                            className="overflow-hidden"
                          >
                            <div className="flex flex-col divide-y divide-black/5">
                              {group.logs.map((log, i) => {
                                const timePart = (log.timestamp || '').split(' - ')[1] || 'N/A';
                                const itemName = items.find(item => item.id === log.id)?.name || log.name || 'Không xác định';
                                
                                return (
                                  <div key={i} className="p-4 hover:bg-black/5 transition-colors flex flex-col gap-3">
                                    <div className="flex justify-between items-start gap-2">
                                      <div className="flex flex-col gap-1.5 flex-1">
                                        <div className="flex items-center gap-2 flex-wrap">
                                          <span className={`inline-flex items-center gap-1 text-[10px] font-bold uppercase tracking-widest px-2 py-0.5 rounded-md ${
                                            log.type === 'Nhập kho' ? 'bg-apple-green/10 text-apple-green' : 
                                            log.type === 'Xuất kho' ? 'bg-apple-blue/10 text-apple-blue' : 
                                            'bg-orange-500/10 text-orange-500'
                                          }`}>
                                            {log.type === 'Nhập kho' && <Plus size={10} />}
                                            {log.type === 'Xuất kho' && <ArrowUpCircle size={10} />}
                                            {log.type === 'Kiểm kê' && <ClipboardList size={10} />}
                                            {log.type}
                                          </span>
                                          <span className="text-[10px] font-bold text-apple-gray">{timePart}</span>
                                        </div>
                                        <p className="text-sm font-bold text-[#1C1C1E]">{itemName}</p>
                                      </div>
                                      <div className="text-right shrink-0">
                                        <p className={`text-base font-bold ${
                                          log.type === 'Nhập kho' ? 'text-apple-green' : 
                                          log.type === 'Xuất kho' ? 'text-apple-blue' : 
                                          'text-[#1C1C1E]'
                                        }`}>
                                          {log.type === 'Nhập kho' ? `+${log.amount || 0}` : 
                                           log.type === 'Xuất kho' ? `-${log.amount || 0}` : 
                                           `${log.actualStock ?? '?'} / ${log.theoreticalStock ?? '?'}`}
                                        </p>
                                        <p className="text-[10px] text-apple-gray font-bold uppercase tracking-widest mt-0.5">{log.unit || ''}</p>
                                      </div>
                                    </div>
                                    <div className="flex items-center gap-1.5 text-[10px] text-apple-gray font-bold uppercase tracking-widest bg-black/5 w-fit px-2.5 py-1 rounded-lg">
                                      {log.user || 'Unknown'}
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </motion.div>
                        )}
                      </AnimatePresence>
                    </div>
                  ))}
                  {groupedLogsForReport.length === 0 && (
                    <div className="text-center py-12">
                      <History size={40} className="mx-auto text-apple-gray/20 mb-4" />
                      <p className="text-apple-gray text-sm font-medium">Không có dữ liệu trong khoảng thời gian này</p>
                    </div>
                  )}
                </div>
              </section>
            </motion.div>
          )}

          {currentScreen === 'add' && (
            <motion.div 
              key="add"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8 pb-24"
            >
              <section className="px-2">
                <span className="text-[10px] font-bold tracking-widest uppercase text-apple-gray mb-1 block">Tạo mới</span>
                <h2 className="text-3xl font-bold text-[#1C1C1E]">Thêm nguyên liệu</h2>
              </section>

              <div className="grid grid-cols-1 gap-6 mt-6">
                <div className="apple-card p-6 space-y-6">
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Tên nguyên vật liệu</label>
                    <input 
                      className="w-full bg-black/5 border-none rounded-2xl h-14 px-5 text-[#1C1C1E] text-lg font-bold focus:ring-0 transition-all" 
                      type="text" 
                      placeholder="Nhập tên nguyên liệu..."
                      value={newItem.name}
                      onChange={(e) => setNewItem(prev => ({ ...prev, name: e.target.value }))}
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Mô tả chi tiết</label>
                    <textarea 
                      className="w-full bg-black/5 border-none rounded-2xl p-5 text-[#1C1C1E] text-base font-medium focus:ring-0 transition-all min-h-[120px]" 
                      placeholder="Nhập mô tả về công dụng, cách bảo quản..."
                      value={newItem.description}
                      onChange={(e) => setNewItem(prev => ({ ...prev, description: e.target.value }))}
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="apple-card p-6 space-y-4">
                    <div className="flex items-center gap-3 mb-2">
                      <div className="w-8 h-8 rounded-xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                        <LayoutGrid size={18} />
                      </div>
                      <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray">Nhóm danh mục</label>
                    </div>
                    <div className="relative">
                      <select 
                        className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0 appearance-none"
                        value={newItem.category}
                        onChange={(e) => setNewItem(prev => ({ ...prev, category: e.target.value as any }))}
                      >
                        {customCategories.map(cat => (
                          <option key={cat} value={cat}>{cat}</option>
                        ))}
                      </select>
                      <ChevronDown className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-apple-gray" />
                    </div>
                  </div>

                  <div className="apple-card p-6 space-y-4">
                    <div className="flex items-center gap-3 mb-2">
                      <div className="w-8 h-8 rounded-xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                        <BarChart3 size={18} />
                      </div>
                      <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray">Đơn vị tính</label>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {customUnits.map((u) => (
                        <button 
                          key={u}
                          type="button"
                          onClick={() => setNewItem(prev => ({ ...prev, unit: u }))}
                          className={`px-4 py-3 rounded-xl text-[10px] font-bold transition-all active:scale-95 ${newItem.unit === u ? 'bg-apple-blue text-white shadow-lg shadow-apple-blue/20' : 'bg-black/5 text-apple-gray hover:bg-black/10'}`}
                        >
                          {u}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="apple-card p-6 space-y-6">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                      <Settings size={20} />
                    </div>
                    <div>
                      <h3 className="text-lg font-bold text-[#1C1C1E]">Thông số lưu kho</h3>
                      <p className="text-apple-gray text-xs font-medium">Thiết lập các ngưỡng tồn kho để nhận cảnh báo kịp thời.</p>
                    </div>
                  </div>
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Tồn kho hiện tại</label>
                        <input 
                          className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0" 
                          type="number" 
                          placeholder="0"
                          value={newItem.actualStock === 0 ? '' : (newItem.actualStock || '')}
                          onChange={(e) => {
                            const val = e.target.value;
                            setNewItem(prev => ({ ...prev, actualStock: val === '' ? 0 : Number(val) }));
                          }}
                        />
                      </div>
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Vị trí kệ</label>
                        <input 
                          className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0" 
                          type="text" 
                          placeholder="Ví dụ: A-12"
                          value={newItem.location || ''}
                          onChange={(e) => setNewItem(prev => ({ ...prev, location: e.target.value }))}
                        />
                      </div>
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Ngưỡng tối thiểu</label>
                        <input 
                          className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0" 
                          type="number" 
                          placeholder="0"
                          value={newItem.minThreshold === 0 ? '' : (newItem.minThreshold || '')}
                          onChange={(e) => {
                            const val = e.target.value;
                            setNewItem(prev => ({ ...prev, minThreshold: val === '' ? 0 : Number(val) }));
                          }}
                        />
                      </div>
                    </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="space-y-2">
                      <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Ngưỡng tối đa</label>
                      <input 
                        className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0" 
                        type="number" 
                        placeholder="0"
                        value={newItem.maxThreshold === 0 ? '' : (newItem.maxThreshold || '')}
                        onChange={(e) => {
                          const val = e.target.value;
                          setNewItem(prev => ({ ...prev, maxThreshold: val === '' ? 0 : Number(val) }));
                        }}
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Giá nhập (VND / {newItem.unit})</label>
                      <input 
                        className="w-full bg-black/5 border-none rounded-2xl h-14 px-4 text-[#1C1C1E] font-bold focus:ring-0" 
                        type="number" 
                        placeholder="0"
                        value={newItem.importPrice === 0 ? '' : (newItem.importPrice || '')}
                        onChange={(e) => {
                          const val = e.target.value;
                          const numVal = val === '' ? 0 : Number(val);
                          const normalized = (newItem.unit === 'Kg' || newItem.unit === 'Lít') ? numVal / 1000 : numVal;
                          setNewItem(prev => ({ ...prev, importPrice: normalized }));
                        }}
                      />
                    </div>
                  </div>
                </div>
              </div>

              <div className="flex gap-4 pt-4">
                <button onClick={() => navigateTo('inventory')} className="flex-1 apple-button-secondary">Hủy bỏ</button>
                <button 
                  onClick={handlePreAddItem}
                  disabled={!newItem.name}
                  className="flex-[2] apple-button-primary flex items-center justify-center gap-2"
                >
                  <PlusCircle className="w-5 h-5" />
                  Tạo nguyên liệu
                </button>
              </div>
            </motion.div>
          )}

          {currentScreen === 'financial' && currentUser?.role === 'quản lí' && (
            <motion.div 
              key="financial"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8 pb-24"
            >
              <section className="px-2 flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div className="flex-1">
                  <h2 className="text-2xl font-bold text-[#1C1C1E]">Tài chính</h2>
                  <p className="text-apple-gray text-xs font-medium mt-1">Giá trị dòng vốn lưu động</p>
                </div>
                <div className="flex-shrink-0 flex flex-wrap items-center gap-2 bg-black/5 p-2 rounded-2xl border border-black/5 relative z-30">
                  <div className="flex items-center gap-1 flex-wrap">
                    <input 
                      type="date" 
                      value={financialStartDate}
                      onChange={(e) => setFinancialStartDate(e.target.value)}
                      className="bg-transparent border-none text-xs font-bold text-[#1C1C1E] focus:ring-0 p-1 cursor-pointer min-w-[130px] relative z-40"
                    />
                    <span className="text-apple-gray text-[10px] font-bold">→</span>
                    <input 
                      type="date" 
                      value={financialEndDate}
                      onChange={(e) => setFinancialEndDate(e.target.value)}
                      className="bg-transparent border-none text-xs font-bold text-[#1C1C1E] focus:ring-0 p-1 cursor-pointer min-w-[130px] relative z-40"
                    />
                  </div>
                  <button 
                    onClick={() => {
                      setActiveFinancialStartDate(financialStartDate);
                      setActiveFinancialEndDate(financialEndDate);
                    }}
                    className="bg-apple-blue text-white px-4 py-2 rounded-xl text-xs font-bold shadow-lg shadow-apple-blue/20 active:scale-95 transition-all flex items-center gap-2"
                  >
                    <Search size={14} />
                    Truy vấn
                  </button>
                </div>
              </section>

              <section className="relative overflow-hidden rounded-[40px] bg-apple-blue p-10 text-white shadow-2xl shadow-apple-blue/30">
                {/* Decoration */}
                <div className="absolute -top-24 -right-24 w-64 h-64 bg-white/10 rounded-full blur-3xl" />
                <div className="absolute -bottom-24 -left-24 w-64 h-64 bg-black/10 rounded-full blur-3xl" />
                
                <div className="relative z-10 space-y-8">
                  <div className="flex items-center gap-3">
                    <div className="w-12 h-12 bg-white/20 backdrop-blur-md rounded-2xl flex items-center justify-center border border-white/20">
                      <DollarSign className="w-6 h-6 text-white" />
                    </div>
                    <span className="text-[10px] font-bold tracking-[0.2em] uppercase text-white/70">Tổng giá trị kho hàng</span>
                  </div>
                  
                  <div className="space-y-2">
                    <h2 className="text-6xl font-black tracking-tighter leading-none">
                      {totalInventoryValue.toLocaleString('vi-VN')}
                      <span className="text-2xl ml-2 text-white/50">VND</span>
                    </h2>
                    <p className="text-white/60 text-sm font-medium">Dòng tiền đang nằm trong nguyên vật liệu</p>
                  </div>

                  <div className="grid grid-cols-2 gap-4 pt-4">
                    <div className="bg-white/10 backdrop-blur-md rounded-3xl p-5 border border-white/10">
                      <p className="text-[10px] font-bold uppercase tracking-widest text-white/60 mb-2">Số lượng mặt hàng</p>
                      <p className="text-2xl font-black">{items.length}</p>
                    </div>
                    <div className="bg-white/10 backdrop-blur-md rounded-3xl p-5 border border-white/10">
                      <p className="text-[10px] font-bold uppercase tracking-widest text-white/60 mb-2">Tổng khối lượng</p>
                      <p className="text-2xl font-black">{totalStockWeight} kg</p>
                    </div>
                  </div>
                </div>
              </section>

              <section className="apple-card p-8">
                <div className="flex items-center justify-between mb-8">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                      <DollarSign size={20} />
                    </div>
                    <h3 className="text-[#1C1C1E] font-bold text-xl">Nhập kho trong kỳ</h3>
                  </div>
                  <div className="text-right">
                    <p className="text-xl font-black text-apple-blue">{totalInboundValueInRange.toLocaleString('vi-VN')}đ</p>
                    <p className="text-[10px] font-bold text-apple-gray uppercase tracking-widest">Tổng chi phí nhập</p>
                  </div>
                </div>

                <div className="space-y-4 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                  {filteredLogsForFinancial.filter(l => l.type === 'Nhập kho').map((log, i) => (
                    <div key={i} className="bg-black/5 rounded-2xl p-4 flex items-center justify-between">
                      <div className="flex items-center gap-4">
                        <div className="w-10 h-10 rounded-xl bg-apple-green/10 text-apple-green flex items-center justify-center">
                          <Plus size={18} />
                        </div>
                        <div>
                          <h4 className="font-bold text-[#1C1C1E] text-sm">{items.find(item => item.id === log.id)?.name || log.name || 'Không xác định'}</h4>
                          <p className="text-[10px] text-apple-gray font-medium">{log.user} • {log.timestamp}</p>
                        </div>
                      </div>
                      <div className="text-right">
                        <p className="text-sm font-bold text-[#1C1C1E]">
                          {((log.amount || 0) * (log.importPrice || 0)).toLocaleString('vi-VN')}đ
                        </p>
                        <p className="text-[10px] text-apple-gray font-bold uppercase tracking-widest">
                          {log.amount} {log.unit} x {log.importPrice?.toLocaleString('vi-VN')}đ
                        </p>
                      </div>
                    </div>
                  ))}
                  {filteredLogsForFinancial.filter(l => l.type === 'Nhập kho').length === 0 && (
                    <div className="text-center py-12">
                      <DollarSign size={40} className="mx-auto text-apple-gray/20 mb-4" />
                      <p className="text-apple-gray text-sm font-medium">Không có dữ liệu nhập kho trong kỳ</p>
                    </div>
                  )}
                </div>
              </section>

              <section className="space-y-6">
                <div className="flex items-center justify-between px-2">
                  <h3 className="text-2xl font-bold text-[#1C1C1E]">Chi tiết tài chính</h3>
                  <span className="px-4 py-1.5 bg-apple-blue/10 text-apple-blue text-[10px] font-bold rounded-full uppercase tracking-widest">Sắp xếp theo giá trị</span>
                </div>

                <div className="space-y-4">
                  {items.sort((a, b) => ((b.actualStock * (b.importPrice || 0)) - (a.actualStock * (a.importPrice || 0)))).map(item => (
                    <div key={item.id} className="apple-card p-5 flex items-center gap-5 active:scale-[0.98] transition-all">
                      <div className="flex-1">
                        <h4 className="font-bold text-[#1C1C1E] text-lg leading-tight">{item.name}</h4>
                        <div className="flex items-center gap-2 mt-1">
                          <span className="text-[10px] font-bold text-apple-gray uppercase tracking-widest">{item.actualStock} {item.unit}</span>
                          <div className="w-1 h-1 rounded-full bg-black/10" />
                          <span className="text-[10px] font-bold text-apple-blue uppercase tracking-widest">
                            {(item.importPrice || 0).toLocaleString('vi-VN')} đ / {item.unit === 'Kg' || item.unit === 'Lít' ? item.unit : 'đv'}
                          </span>
                        </div>
                      </div>
                      <div className="text-right">
                        <p className="text-lg font-black text-[#1C1C1E]">
                          {(item.actualStock * (item.importPrice || 0)).toLocaleString('vi-VN')}đ
                        </p>
                        <p className="text-[10px] font-bold text-apple-gray uppercase tracking-widest opacity-60">Giá trị tồn</p>
                      </div>
                    </div>
                  ))}
                </div>
              </section>
            </motion.div>
          )}

          {currentScreen === 'suppliers' && currentUser?.role === 'quản lí' && (
            <motion.div 
              key="suppliers"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8 pb-24"
            >
              <section className="px-2 flex items-center justify-between">
                <div>
                  <h2 className="text-2xl font-bold text-[#1C1C1E]">Mua hàng</h2>
                  <p className="text-apple-gray text-xs font-medium mt-1">Quản lý nhà cung cấp & So sánh giá</p>
                </div>
                <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                  <ShoppingCart size={20} />
                </div>
              </section>

              <div className="space-y-6">
                {purchaseCart.length > 0 && (
                  <div className="bg-white rounded-3xl p-6 shadow-xl shadow-apple-blue/5 border border-apple-blue/10 space-y-4">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                          <ShoppingCart size={20} />
                        </div>
                        <div>
                          <h3 className="font-bold text-[#1C1C1E] text-lg">Giỏ hàng ({purchaseCart.length})</h3>
                          <p className="text-[10px] font-bold text-apple-gray uppercase tracking-widest">Dự kiến đặt hàng</p>
                        </div>
                      </div>
                      <button onClick={() => setPurchaseCart([])} className="text-xs font-bold text-red-500 bg-red-50 px-3 py-1.5 rounded-lg active:scale-95 transition-all">
                        Xóa tất cả
                      </button>
                    </div>
                    <div className="space-y-3">
                      {purchaseCart.map((cartItem, idx) => {
                        const item = items.find(i => i.id === cartItem.itemId);
                        return (
                          <div key={idx} className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 p-3 bg-black/5 rounded-2xl">
                            <div>
                              <p className="font-bold text-sm text-[#1C1C1E]">{item?.name || 'Không xác định'} <span className="font-medium opacity-60">x{cartItem.amount} {item?.unit}</span></p>
                              <p className="text-xs font-medium text-apple-gray mt-0.5">{(cartItem.price * cartItem.amount).toLocaleString('vi-VN')} đ (Dự kiến: {cartItem.price.toLocaleString('vi-VN')}đ/{item?.unit})</p>
                            </div>
                            <button onClick={() => handleRemovePurchaseCart(cartItem.itemId)} className="w-8 h-8 rounded-lg bg-red-100 flex items-center justify-center text-red-500 ml-auto flex-shrink-0">
                              <Minus size={16} />
                            </button>
                          </div>
                        );
                      })}
                    </div>
                    <div className="pt-4 border-t border-black/5 flex items-center justify-between">
                      <p className="font-bold text-sm text-apple-gray">Tổng dự kiến:</p>
                      <p className="font-black text-xl text-apple-blue">
                        {purchaseCart.reduce((sum, item) => sum + item.price * item.amount, 0).toLocaleString('vi-VN')} đ
                      </p>
                    </div>
                    <button 
                      onClick={() => {
                        const newInboundCart = purchaseCart.map(p => ({
                          itemId: p.itemId,
                          amount: p.amount,
                          reason: 'Mới mua hàng'
                        }));
                        setInboundCart(prev => [...prev, ...newInboundCart]);
                        setPurchaseCart([]);
                        navigateTo('inbound');
                      }}
                      className="w-full h-12 bg-apple-blue font-bold text-white rounded-xl flex items-center justify-center gap-2 active:scale-95 transition-all mt-4 hover:shadow-lg hover:shadow-apple-blue/30"
                    >
                      <ArrowDownCircle size={18} />
                      Chuyển sang Nhập kho & Đặt hàng
                    </button>
                  </div>
                )}
                {items.map(item => (
                  <div key={item.id} className="apple-card p-6 space-y-5">
                    <div className="flex flex-col gap-4">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-4">
                          <div>
                            <h3 className="font-bold text-[#1C1C1E] text-lg">{item.name}</h3>
                            <p className="text-[10px] font-bold text-apple-gray uppercase tracking-widest">Giá hiện tại: {item.importPrice?.toLocaleString('vi-VN')} đ / {item.unit}</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2">
                          <button 
                            onClick={() => {
                              setSelectedItemId(item.id);
                              setEditingSupplier({ name: '', contact: '', price: 0 });
                              setShowSupplierModal(true);
                            }}
                            className="w-10 h-10 bg-black/5 text-apple-gray rounded-xl flex items-center justify-center hover:bg-black/10 hover:text-[#1C1C1E] active:scale-90 transition-all font-bold group relative"
                            title="Thêm nhà cung cấp"
                          >
                            <Plus size={20} />
                          </button>
                        </div>
                      </div>
                      <div className="flex items-center gap-3">
                        <div className="relative flex-1">
                          <input 
                            type="number" 
                            placeholder={`Số lượng cần nhập (${item.unit})`}
                            value={purchaseInput[item.id] || ''}
                            onChange={(e) => setPurchaseInput(prev => ({...prev, [item.id]: e.target.value}))}
                            className="w-full h-12 bg-black/5 border-none rounded-xl px-4 text-sm font-bold text-[#1C1C1E] focus:ring-2 focus:ring-apple-blue/20"
                          />
                        </div>
                        <button 
                          onClick={() => handleAddPurchaseCart(item)}
                          disabled={!purchaseInput[item.id] || Number(purchaseInput[item.id]) <= 0}
                          className="h-12 px-6 bg-apple-blue text-white rounded-xl font-bold flex items-center justify-center gap-2 active:scale-95 transition-all disabled:opacity-50 disabled:active:scale-100"
                        >
                          <ShoppingCart size={18} />
                          Thêm
                        </button>
                      </div>
                    </div>

                    {item.suppliers && item.suppliers.length > 0 ? (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                        {item.suppliers.map((s, idx) => (
                          <div key={idx} className="bg-black/5 rounded-2xl p-4 flex items-center justify-between group relative">
                            <div className="flex-1">
                              <div className="flex items-center gap-2">
                                <p className="font-bold text-sm text-[#1C1C1E]">{s.name}</p>
                                <span className="text-[8px] font-bold text-apple-gray uppercase tracking-tighter bg-white px-1.5 py-0.5 rounded border border-black/5">{s.lastUpdated?.split(' - ')[0]}</span>
                              </div>
                              <p className="text-[10px] text-apple-gray font-medium mt-0.5">{s.contact}</p>
                            </div>
                            <div className="text-right">
                              <p className={`font-bold text-sm ${s.price > (item.importPrice || 0) ? 'text-red-500' : 'text-apple-green'}`}>
                                {s.price.toLocaleString('vi-VN')}đ
                              </p>
                              {s.price > (item.importPrice || 0) && (
                                <p className="text-[8px] font-black text-red-500 uppercase tracking-widest flex items-center gap-1 justify-end mt-0.5">
                                  <AlertTriangle size={8} /> Tăng giá
                                </p>
                              )}
                            </div>
                            <div className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity flex gap-1">
                              <button 
                                onClick={() => {
                                  setSelectedItemId(item.id);
                                  setEditingSupplier({ ...s, idx });
                                  setShowSupplierModal(true);
                                }}
                                className="p-1.5 bg-white rounded-lg shadow-sm text-apple-gray hover:text-apple-blue"
                              >
                                <Edit3 size={12} />
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="bg-black/5 rounded-2xl p-6 text-center">
                        <p className="text-apple-gray text-xs font-medium italic">Chưa có thông tin nhà cung cấp cho mặt hàng này</p>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </motion.div>
          )}

          {currentScreen === 'inventory' && (
            <motion.div 
              key="inventory"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 1.05 }}
              className="space-y-6 pb-24"
            >
              <section className="flex items-center justify-between px-2">
                <div>
                  <h2 className="text-2xl font-bold text-[#1C1C1E]">Danh mục</h2>
                  <p className="text-apple-gray text-xs font-medium mt-1">Quản lý tất cả dược liệu</p>
                </div>
                <div className="flex gap-2 relative z-50">
                  <button 
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      navigateTo('add');
                    }}
                    className="bg-apple-blue text-white p-2.5 rounded-full shadow-lg shadow-apple-blue/20 active:scale-90 transition-all hover:bg-apple-blue/90 flex items-center justify-center cursor-pointer"
                    title="Thêm mới"
                  >
                    <Plus size={20} />
                  </button>
                </div>
              </section>

              {/* Search and Filter in Inventory */}
              <div className="space-y-4">
                <div className="relative group">
                  <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-apple-gray group-focus-within:text-apple-blue transition-colors" />
                  <input 
                    type="text"
                    placeholder="Tìm kiếm dược liệu..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full bg-black/5 border-none rounded-2xl py-4 pl-12 pr-4 text-[#1C1C1E] focus:ring-2 focus:ring-apple-blue/20 transition-all font-medium"
                  />
                  {searchQuery && (
                    <button 
                      onClick={() => setSearchQuery('')}
                      className="absolute right-4 top-1/2 -translate-y-1/2 p-1 hover:bg-black/10 rounded-full transition-all"
                    >
                      <X size={16} className="text-apple-gray" />
                    </button>
                  )}
                  {searchSuggestions.length > 0 && (
                    <div className="absolute top-full left-0 right-0 mt-2 bg-white/80 backdrop-blur-xl rounded-2xl shadow-2xl z-[60] overflow-hidden border border-black/5">
                      {searchSuggestions.map(item => (
                        <button 
                          key={item.id}
                          onClick={() => {
                            setSelectedItemId(item.id);
                            navigateTo('details');
                            setSearchQuery('');
                          }}
                          className="w-full px-6 py-4 flex items-center gap-4 hover:bg-apple-blue/5 transition-colors text-left border-b border-black/5 last:border-0"
                        >
                          <div className="flex-1">
                            <p className="text-[#1C1C1E] font-bold text-sm">{item.name}</p>
                            <p className="text-apple-gray text-[10px] font-bold uppercase tracking-widest">{item.category}</p>
                          </div>
                          <ArrowRight size={16} className="text-apple-gray" />
                        </button>
                      ))}
                    </div>
                  )}
                </div>

                <div className="flex gap-2 overflow-x-auto no-scrollbar pb-1 px-1">
                  {[
                    { id: 'all', label: 'Tất cả' },
                    { id: 'low_stock', label: 'Sắp hết hàng', icon: AlertTriangle },
                    { id: 'urgent', label: 'Cần nhập gấp', icon: AlertOctagon },
                    { id: 'expiring', label: 'Sắp hết hạn', icon: Clock }
                  ].map((filter) => (
                    <button
                      key={filter.id}
                      onClick={() => setInventoryQuickFilter(filter.id as any)}
                      className={`whitespace-nowrap px-4 py-2 rounded-xl text-xs font-bold transition-all active:scale-95 border flex items-center gap-1.5 ${
                        inventoryQuickFilter === filter.id
                          ? 'bg-[#1C1C1E] text-white shadow-lg border-[#1C1C1E]'
                          : 'bg-white text-[#1C1C1E] border-black/10 hover:bg-black/5'
                      }`}
                    >
                      {filter.icon && <filter.icon size={14} className={inventoryQuickFilter === filter.id ? 'text-white' : filter.id === 'low_stock' ? 'text-orange-500' : filter.id === 'urgent' ? 'text-red-500' : 'text-blue-500'} />}
                      {filter.label}
                    </button>
                  ))}
                </div>

                <div className="flex gap-2 overflow-x-auto no-scrollbar pb-1 px-1">
                  {['Tất cả', ...customCategories].map((cat) => (
                    <button
                      key={cat}
                      onClick={() => setInventoryCategory(cat)}
                      className={`whitespace-nowrap px-5 py-2.5 rounded-xl text-xs font-bold transition-all active:scale-95 border ${
                        inventoryCategory === cat
                          ? 'bg-apple-blue text-white shadow-lg shadow-apple-blue/20 border-apple-blue'
                          : 'bg-white text-apple-gray border-black/5 hover:bg-black/5'
                      }`}
                    >
                      {cat}
                    </button>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-1 gap-3">
                {filteredItems.length > 0 ? (
                  filteredItems.map((item) => (
                    <SwipeableLogItem
                      key={item.id}
                      onEdit={() => navigateTo('edit', item.id)}
                      onDelete={() => {
                        setSelectedItemId(item.id);
                        setConfirmConfig({
                          title: 'Xóa dược liệu',
                          message: `Bạn có chắc chắn muốn xóa "${item.name}" khỏi hệ thống không?`,
                          type: 'danger',
                          icon: <Trash2 className="w-10 h-10" />,
                          onConfirm: () => handleDeleteItem(item.id)
                        });
                        setShowConfirmModal(true);
                      }}
                    >
                      <div 
                        onClick={() => navigateTo('details', item.id)}
                        className="apple-card p-5 flex items-center justify-between border-transparent hover:border-apple-blue/20 transition-all duration-300 cursor-pointer active:scale-[0.99]"
                      >
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 mb-2">
                            <span className="text-[10px] font-bold uppercase tracking-widest text-apple-blue px-2 py-0.5 bg-apple-blue/5 rounded-md">{item.category}</span>
                            <span className="text-[10px] font-bold uppercase tracking-widest text-[#1C1C1E] px-2 py-0.5 bg-black/5 rounded-md">{item.unit}</span>
                            {item.actualStock < (item.minThreshold || 0) && (
                              <span className="text-[10px] font-bold uppercase tracking-widest text-red-500 px-2 py-0.5 bg-red-50 rounded-md">Tồn thấp</span>
                            )}
                          </div>
                          <h3 className="text-lg font-extrabold text-[#1C1C1E] mb-2 truncate group-hover:text-apple-blue transition-colors">{item.name}</h3>
                          <div className="flex items-center gap-6">
                            <div className="flex flex-col">
                              <span className="text-[10px] font-bold text-apple-gray uppercase tracking-widest mb-0.5">Tồn kho</span>
                              <span className={`text-sm font-black ${item.actualStock < (item.minThreshold || 0) ? 'text-red-500' : 'text-[#1C1C1E]'}`}>
                                {item.actualStock} {item.unit}
                              </span>
                            </div>
                            <div className="flex flex-col">
                              <span className="text-[10px] font-bold text-apple-gray uppercase tracking-widest mb-0.5">Vị trí</span>
                              <span className="text-sm font-bold text-apple-gray">{item.location || '—'}</span>
                            </div>
                            <div className="flex flex-col">
                              <span className="text-[10px] font-bold text-apple-gray uppercase tracking-widest mb-0.5">Giá nhập</span>
                              <span className="text-sm font-bold text-apple-gray">{(item.importPrice || 0).toLocaleString()}đ</span>
                            </div>
                          </div>
                        </div>
                        
                        <div className="flex items-center gap-4 pl-4 border-l border-black/5">
                          <div className="hidden md:flex flex-col gap-2">
                            <button 
                              onClick={(e) => {
                                e.stopPropagation();
                                navigateTo('edit', item.id);
                              }}
                              className="p-2.5 bg-apple-blue/10 text-apple-blue rounded-xl active:scale-90 transition-all hover:bg-apple-blue hover:text-white"
                            >
                              <Edit3 size={18} />
                            </button>
                            <button 
                              onClick={(e) => {
                                e.stopPropagation();
                                setSelectedItemId(item.id);
                                setConfirmConfig({
                                  title: 'Xóa dược liệu',
                                  message: `Bạn có chắc chắn muốn xóa "${item.name}" khỏi hệ thống không?`,
                                  type: 'danger',
                                  icon: <Trash2 className="w-10 h-10" />,
                                  onConfirm: () => handleDeleteItem(item.id)
                                });
                                setShowConfirmModal(true);
                              }}
                              className="p-2.5 bg-red-500/10 text-red-500 rounded-xl active:scale-90 transition-all hover:bg-red-500 hover:text-white"
                            >
                              <Trash2 size={18} />
                            </button>
                          </div>
                          <ChevronRight className="w-5 h-5 text-apple-gray/20" />
                        </div>
                      </div>
                    </SwipeableLogItem>
                  ))
                ) : (
                  <div className="p-12 text-center apple-card flex flex-col items-center justify-center border-dashed border-2 border-black/5">
                    <div className="w-16 h-16 bg-black/5 rounded-full flex items-center justify-center mb-4">
                      <Search className="w-8 h-8 text-apple-gray" />
                    </div>
                    <p className="text-[#1C1C1E] font-bold">Không tìm thấy kết quả</p>
                    <p className="text-apple-gray text-xs mt-1">Vui lòng thử từ khóa hoặc danh mục khác</p>
                  </div>
                )}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
        </PullToRefresh>
      </main>

      {/* Bottom Navigation Bar - Removed in favor of Sidebar */}

      {/* Confirmation Modal */}
      <AnimatePresence>
        {showConfirmModal && (
          <div className="fixed inset-0 bg-black/20 backdrop-blur-sm z-[100] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-sm apple-card p-8 flex flex-col items-center text-center"
            >
              <div className={`mb-6 w-20 h-20 rounded-full flex items-center justify-center ${confirmConfig?.type === 'danger' ? 'bg-red-500/10 text-red-500' : 'bg-apple-blue/10 text-apple-blue'}`}>
                {confirmConfig?.icon || <Save className="w-10 h-10" />}
              </div>
              
              <h2 className="text-xl font-bold text-[#1C1C1E] mb-2">{confirmConfig?.title || 'Xác nhận'}</h2>
              <p className="text-apple-gray text-sm mb-8">{confirmConfig?.message || 'Bạn có chắc chắn muốn thực hiện thao tác này không?'}</p>
              
              <div className="flex flex-col gap-3 w-full">
                <button 
                  onClick={() => {
                    if (confirmConfig?.onConfirm) {
                      confirmConfig.onConfirm();
                    } else {
                      handleAuditConfirm();
                    }
                  }}
                  disabled={isSyncing}
                  className={`w-full apple-button-primary flex items-center justify-center gap-2 ${confirmConfig?.type === 'danger' ? 'bg-red-500 shadow-red-500/20' : ''}`}
                >
                  {isSyncing ? <IOSSpinner /> : (
                    <>
                      <CheckCircle2 className="w-4 h-4" />
                      <span>Xác nhận</span>
                    </>
                  )}
                </button>
                <button 
                  onClick={() => {
                    setShowConfirmModal(false);
                    setConfirmConfig(null);
                  }}
                  className="w-full apple-button-secondary"
                >
                  Hủy bỏ
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Supplier Modal */}
      <AnimatePresence>
        {showSupplierModal && (
          <div className="fixed inset-0 bg-black/20 backdrop-blur-sm z-[100] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-md apple-card p-8 space-y-6"
            >
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center">
                  <ShoppingCart size={20} />
                </div>
                <div>
                  <h2 className="text-xl font-bold text-[#1C1C1E]">{editingSupplier.idx !== undefined ? 'Sửa nhà cung cấp' : 'Thêm nhà cung cấp'}</h2>
                  <p className="text-apple-gray text-[10px] font-bold uppercase tracking-widest">{selectedItem?.name}</p>
                </div>
              </div>

              <div className="space-y-4">
                <div className="space-y-2">
                  <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Tên nhà cung cấp</label>
                  <input 
                    className="w-full bg-black/5 border-none rounded-2xl h-14 px-5 text-[#1C1C1E] font-bold focus:ring-0" 
                    type="text" 
                    placeholder="Ví dụ: Công ty Dược ABC"
                    value={editingSupplier.name}
                    onChange={(e) => setEditingSupplier(prev => ({ ...prev, name: e.target.value }))}
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Thông tin liên hệ</label>
                  <input 
                    className="w-full bg-black/5 border-none rounded-2xl h-14 px-5 text-[#1C1C1E] font-bold focus:ring-0" 
                    type="text" 
                    placeholder="Số điện thoại hoặc địa chỉ..."
                    value={editingSupplier.contact}
                    onChange={(e) => setEditingSupplier(prev => ({ ...prev, contact: e.target.value }))}
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-bold uppercase tracking-widest text-apple-gray ml-2">Giá cung cấp (VND / {selectedItem?.unit})</label>
                  <input 
                    className="w-full bg-black/5 border-none rounded-2xl h-14 px-5 text-[#1C1C1E] font-bold focus:ring-0" 
                    type="number" 
                    placeholder="0"
                    value={editingSupplier.price === 0 ? '' : (editingSupplier.price || '')}
                    onChange={(e) => {
                      const val = e.target.value;
                      setEditingSupplier(prev => ({ ...prev, price: val === '' ? 0 : Number(val) }));
                    }}
                  />
                </div>
              </div>

              <div className="flex gap-3 pt-4">
                <button 
                  onClick={() => setShowSupplierModal(false)}
                  className="flex-1 apple-button-secondary"
                >
                  Hủy bỏ
                </button>
                <button 
                  onClick={handleSaveSupplier}
                  disabled={isSyncing || !editingSupplier.name}
                  className="flex-[2] apple-button-primary flex items-center justify-center gap-2"
                >
                  {isSyncing ? <IOSSpinner /> : (
                    <>
                      <CheckCircle2 className="w-4 h-4" />
                      <span>Lưu thông tin</span>
                    </>
                  )}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Delete All Data Password Modal */}
      <AnimatePresence>
        {showDeletePasswordModal && (
          <div className="fixed inset-0 bg-black/30 backdrop-blur-sm z-[110] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-sm apple-card p-8 flex flex-col items-center text-center"
            >
              <div className="mb-6 w-16 h-16 bg-red-500/10 rounded-full flex items-center justify-center text-red-500">
                <ShieldCheck className="w-8 h-8" />
              </div>
              
              <h2 className="text-xl font-bold text-[#1C1C1E] mb-2">Bảo mật hệ thống</h2>
              <p className="text-apple-gray text-sm mb-6">Vui lòng nhập mật khẩu xác nhận để xóa toàn bộ dữ liệu. Thao tác này không thể khôi phục.</p>
              
              <div className="w-full space-y-6">
                <div className="relative">
                  <input 
                    type="password"
                    placeholder="Mật khẩu"
                    value={deletePasswordInput}
                    onChange={(e) => {
                      setDeletePasswordInput(e.target.value);
                      setPasswordError(false);
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') handleDeleteAllData();
                    }}
                    className={`w-full bg-black/5 border-2 rounded-2xl py-4 px-5 text-center text-xl font-bold tracking-widest focus:ring-0 transition-all ${
                      passwordError ? 'border-red-500 text-red-500' : 'border-transparent focus:border-apple-blue'
                    }`}
                    autoFocus
                  />
                  {passwordError && (
                    <motion.p 
                      initial={{ opacity: 0, y: -10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="text-red-500 text-xs font-bold mt-2"
                    >
                      Mật khẩu không chính xác!
                    </motion.p>
                  )}
                </div>

                <div className="flex flex-col gap-3">
                  <button 
                    onClick={handleDeleteAllData}
                    className="w-full apple-button-primary bg-red-500 shadow-red-500/20"
                  >
                    Xác nhận xóa
                  </button>
                  <button 
                    onClick={() => setShowDeletePasswordModal(false)}
                    className="w-full apple-button-secondary"
                  >
                    Hủy bỏ
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
    </div>
  );
}
