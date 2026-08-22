export type Category = string;
export type Unit = string;
export type AuditFrequency = 'daily' | 'weekly' | 'monthly';
export type Role = 'nhân viên' | 'quản lí';

export interface User {
  uid: string;
  username: string;
  role: Role;
  name: string;
}

export interface Supplier {
  id: string;
  name: string;
  contact: string;
  price: number;
  lastUpdated: string;
}

export interface InventoryItem {
  id: string;
  name: string;
  category: Category;
  unit: string; // Changed to string to be more flexible with PDF units
  batchId?: string;
  expiryDate?: string;
  theoreticalStock?: number;
  actualStock: number;
  location?: string;
  minThreshold?: number;
  maxThreshold?: number;
  status?: 'Ổn định' | 'Sẵn có' | 'Sắp hết' | 'Tồn thấp' | 'Cần nhập gấp';
  description?: string;
  auditFrequency?: AuditFrequency[];
  importPrice?: number; // Price per base unit (e.g., per Gram, per ml, per Cái)
  suppliers?: Supplier[];
}

export interface Transaction {
  id: string;
  itemId: string;
  type: 'Nhập kho' | 'Xuất kho';
  reason: string;
  amount: number;
  unit: Unit;
  timestamp: string;
}
