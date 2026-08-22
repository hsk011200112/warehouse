import { InventoryItem, Transaction } from './types';

export const MOCK_ITEMS: InventoryItem[] = [
  {
    id: '1',
    name: 'Cỏ ngọt',
    category: 'Thảo mộc thô',
    unit: 'Gram',
    actualStock: 18.5,
    status: 'Ổn định',
    location: 'A-12',
    minThreshold: 500,
    description: 'Hàng khô cao cấp',
    auditFrequency: ['weekly', 'monthly']
  },
  {
    id: '2',
    name: 'Thục địa',
    category: 'Thảo mộc thô',
    unit: 'Gram',
    actualStock: 18.5,
    status: 'Sẵn có',
    location: 'B-04',
    minThreshold: 1000,
    batchId: '#45',
    expiryDate: '12/2024',
    description: 'Thành phẩm 9 lần chưng',
    auditFrequency: ['weekly', 'monthly']
  },
  {
    id: '3',
    name: 'Mật ong rừng',
    category: 'Pha chế',
    unit: 'ml',
    actualStock: 1.2,
    status: 'Tồn thấp',
    description: 'Cần nhập gấp',
    auditFrequency: ['weekly', 'monthly']
  },
  {
    id: '4',
    name: 'Trà lài Tây Bắc',
    category: 'Thảo mộc thô',
    unit: 'Kg',
    actualStock: 42,
    status: 'Ổn định',
    location: 'C-09',
    description: 'Hàng nhập vụ đông',
    auditFrequency: ['weekly', 'monthly']
  },
  {
    id: '5',
    name: 'Sâm mộc vị',
    category: 'Pha chế',
    unit: 'ml',
    actualStock: 32,
    batchId: '#SMV-092',
    auditFrequency: ['daily']
  },
  {
    id: '6',
    name: 'Nước dừa tươi',
    category: 'Pha chế',
    unit: 'ml',
    actualStock: 88,
    description: 'Nguồn cung tiêu chuẩn hàng ngày. Kiểm tra niêm phong trước khi đếm.',
    auditFrequency: ['daily']
  },
  {
    id: '7',
    name: 'Chanh dây',
    category: 'Vật tư',
    unit: 'Gram',
    actualStock: 12,
    auditFrequency: ['daily', 'weekly', 'monthly']
  },
  {
    id: '8',
    name: 'Quả tắc',
    category: 'Vật tư',
    unit: 'Gram',
    actualStock: 4.5,
    auditFrequency: ['daily', 'weekly', 'monthly']
  },
  {
    id: '9',
    name: 'Cam sành',
    category: 'Vật tư',
    unit: 'Gram',
    actualStock: 21,
    auditFrequency: ['daily', 'weekly', 'monthly']
  },
  {
    id: '10',
    name: 'Huyền sâm',
    category: 'Thảo mộc thô',
    unit: 'Gram',
    actualStock: 1200,
    status: 'Sẵn có',
    description: 'Rễ phơi khô, chất lượng loại 1',
    auditFrequency: ['weekly', 'monthly']
  },
  {
    id: '11',
    name: 'Thanh mai',
    category: 'Vật tư',
    unit: 'Kg',
    actualStock: 8.2,
    status: 'Ổn định',
    description: 'Quả tươi nhập mới',
    auditFrequency: ['daily', 'weekly', 'monthly']
  },
  {
    id: '12',
    name: 'Sâm hỷ',
    category: 'Pha chế',
    unit: 'ml',
    actualStock: 500,
    status: 'Sẵn có',
    description: 'Thành phẩm pha sẵn',
    auditFrequency: ['daily']
  }
];

export const MOCK_TRANSACTIONS: Transaction[] = [
  {
    id: 't1',
    itemId: '2',
    type: 'Xuất kho',
    reason: 'Bốc thuốc',
    amount: 250,
    unit: 'Gram',
    timestamp: '24/05/2024 • 14:30'
  },
  {
    id: 't2',
    itemId: '2',
    type: 'Nhập kho',
    reason: 'Lô mới',
    amount: 5000,
    unit: 'Gram',
    timestamp: '20/05/2024 • 09:15'
  }
];
