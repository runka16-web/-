import React, { useState, useMemo } from 'react';
import { 
  Truck, 
  Package, 
  Plus, 
  Trash2, 
  Play, 
  BarChart3, 
  LayoutDashboard, 
  Settings,
  ChevronRight,
  AlertCircle,
  CheckCircle2,
  Info,
  FileSpreadsheet,
  Menu,
  X,
  Download,
  Globe,
  ChevronDown
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { Good, Vehicle, OptimizationResult } from './types';
import { optimizeLoading } from './optimizer';
import { Language, translations } from './i18n';

// Default Data
const DEFAULT_GOODS: Good[] = [
  { id: '1', name: '电子产品 (Electronics)', weight: 15, volume: 0.1, quantity: 50 },
  { id: '2', name: '家具组件 (Furniture Parts)', weight: 45, volume: 0.8, quantity: 20 },
  { id: '3', name: '生鲜箱 (Fresh Produce)', weight: 12, volume: 0.05, quantity: 100 },
  { id: '4', name: '机械零件 (Machinery)', weight: 200, volume: 1.5, quantity: 5 },
];

const DEFAULT_VEHICLES: Vehicle[] = [
  { id: 'v1', type: '4.2米轻型厢式货车', maxWeight: 1500, maxVolume: 15, count: 2 },
  { id: 'v2', type: '6.8米中型货车', maxWeight: 8000, maxVolume: 35, count: 1 },
  { id: 'v3', type: '9.6米大型货车', maxWeight: 18000, maxVolume: 55, count: 1 },
];

export default function App() {
  const [goods, setGoods] = useState<Good[]>(DEFAULT_GOODS);
  const [vehicles, setVehicles] = useState<Vehicle[]>(DEFAULT_VEHICLES);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'goods' | 'vehicles' | 'results'>('dashboard');
  const [result, setResult] = useState<OptimizationResult | null>(null);
  const [selectedGoods, setSelectedGoods] = useState<Set<string>>(new Set());
  const [selectedVehicles, setSelectedVehicles] = useState<Set<string>>(new Set());
  const [lang, setLang] = useState<Language>('zh');

  const t = translations[lang];

  // Handlers for Goods
  const addGood = () => {
    const newGood: Good = {
      id: Math.random().toString(36).substr(2, 9),
      name: t.newGood,
      weight: 0,
      volume: 0,
      quantity: 1
    };
    setGoods([...goods, newGood]);
  };

  const updateGood = (id: string, field: keyof Good, value: string | number) => {
    setGoods(goods.map(g => g.id === id ? { ...g, [field]: value } : g));
  };

  const removeGood = (id: string) => {
    setGoods(goods.filter(g => g.id !== id));
    const newSelected = new Set(selectedGoods);
    newSelected.delete(id);
    setSelectedGoods(newSelected);
  };

  const toggleGoodSelection = (id: string) => {
    const newSelected = new Set(selectedGoods);
    if (newSelected.has(id)) {
      newSelected.delete(id);
    } else {
      newSelected.add(id);
    }
    setSelectedGoods(newSelected);
  };

  const toggleAllGoods = () => {
    if (selectedGoods.size === goods.length) {
      setSelectedGoods(new Set());
    } else {
      setSelectedGoods(new Set(goods.map(g => g.id)));
    }
  };

  const deleteSelectedGoods = () => {
    setGoods(goods.filter(g => !selectedGoods.has(g.id)));
    setSelectedGoods(new Set());
  };

  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws) as any[];

      const importedGoods: Good[] = data.map((item) => ({
        id: Math.random().toString(36).substr(2, 9),
        name: item['名称'] || item['Name'] || item['物资名称'] || '未知物品',
        weight: parseFloat(item['重量'] || item['Weight'] || item['单件重量'] || 0),
        volume: parseFloat(item['体积'] || item['Volume'] || item['单件体积'] || 0),
        quantity: parseInt(item['数量'] || item['Quantity'] || 1),
      }));

      setGoods([...goods, ...importedGoods]);
    };
    reader.readAsBinaryString(file);
  };

  const downloadTemplate = () => {
    const templateData = [
      { [t.goodName]: t.exampleGood1, [t.weightPerUnit.split(' ')[0]]: 10, [t.volumePerUnit.split(' ')[0]]: 0.5, [t.quantity]: 100 },
      { [t.goodName]: t.exampleGood2, [t.weightPerUnit.split(' ')[0]]: 25, [t.volumePerUnit.split(' ')[0]]: 1.2, [t.quantity]: 50 },
    ];
    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, t.templateSheetName);
    XLSX.writeFile(wb, t.templateFileName);
  };

  // Handlers for Vehicles
  const addVehicle = () => {
    const newVehicle: Vehicle = {
      id: Math.random().toString(36).substr(2, 9),
      type: t.newVehicle,
      maxWeight: 0,
      maxVolume: 0,
      count: 1
    };
    setVehicles([...vehicles, newVehicle]);
  };

  const updateVehicle = (id: string, field: keyof Vehicle, value: string | number) => {
    setVehicles(vehicles.map(v => v.id === id ? { ...v, [field]: value } : v));
  };

  const removeVehicle = (id: string) => {
    setVehicles(vehicles.filter(v => v.id !== id));
    const newSelected = new Set(selectedVehicles);
    newSelected.delete(id);
    setSelectedVehicles(newSelected);
  };

  const toggleVehicleSelection = (id: string) => {
    const newSelected = new Set(selectedVehicles);
    if (newSelected.has(id)) {
      newSelected.delete(id);
    } else {
      newSelected.add(id);
    }
    setSelectedVehicles(newSelected);
  };

  const toggleAllVehicles = () => {
    if (selectedVehicles.size === vehicles.length) {
      setSelectedVehicles(new Set());
    } else {
      setSelectedVehicles(new Set(vehicles.map(v => v.id)));
    }
  };

  const deleteSelectedVehicles = () => {
    setVehicles(vehicles.filter(v => !selectedVehicles.has(v.id)));
    setSelectedVehicles(new Set());
  };

  const handleOptimize = () => {
    const res = optimizeLoading(goods, vehicles);
    setResult(res);
    setActiveTab('results');
  };

  const deleteLoad = (vehicleId: string) => {
    if (!result) return;
    const newLoads = result.loads.filter(l => l.vehicleId !== vehicleId);
    setResult({
      ...result,
      loads: newLoads,
      summary: {
        ...result.summary,
        totalVehiclesUsed: newLoads.length
        // Note: utilization averages would need recalculation for full accuracy
      }
    });
  };

  const updateLoadItem = (vehicleId: string, goodId: string, newQuantity: number) => {
    if (!result) return;
    const newLoads = result.loads.map(load => {
      if (load.vehicleId !== vehicleId) return load;
      
      const newItems = load.items.map(item => {
        if (item.goodId !== goodId) return item;
        const goodRef = goods.find(g => g.id === item.goodId);
        const weight = (goodRef?.weight || 0) * newQuantity;
        const volume = (goodRef?.volume || 0) * newQuantity;
        return { ...item, quantity: newQuantity, weight, volume };
      });

      const totalWeight = newItems.reduce((sum, i) => sum + i.weight, 0);
      const totalVolume = newItems.reduce((sum, i) => sum + i.volume, 0);
      
      // Find vehicle capacity
      const vehicleType = vehicles.find(v => v.type === load.vehicleType);
      const weightUtilization = vehicleType ? (totalWeight / vehicleType.maxWeight) * 100 : 0;
      const volumeUtilization = vehicleType ? (totalVolume / vehicleType.maxVolume) * 100 : 0;

      return { ...load, items: newItems, totalWeight, totalVolume, weightUtilization, volumeUtilization };
    });

    setResult({ ...result, loads: newLoads });
  };

  return (
    <div className="min-h-screen bg-[#F5F5F5] text-[#1A1A1A] font-sans pb-24 md:pb-0">
      {/* Desktop Sidebar */}
      <aside className="hidden md:block fixed left-0 top-0 h-full w-64 bg-white border-r border-[#E5E5E5] z-20">
        <div className="p-6 border-bottom border-[#E5E5E5]">
          <div className="flex items-center gap-3 text-emerald-600 font-bold text-xl">
            <Truck className="w-8 h-8" />
            <span>{t.appName}</span>
          </div>
          <p className="text-xs text-gray-400 mt-1 uppercase tracking-widest font-semibold">{t.appDesc}</p>
        </div>

        <nav className="mt-6 px-4 space-y-2">
          <NavItem 
            icon={<LayoutDashboard className="w-5 h-5" />} 
            label={t.dashboard} 
            active={activeTab === 'dashboard'} 
            onClick={() => setActiveTab('dashboard')} 
          />
          <NavItem 
            icon={<Package className="w-5 h-5" />} 
            label={t.goods} 
            active={activeTab === 'goods'} 
            onClick={() => setActiveTab('goods')} 
          />
          <NavItem 
            icon={<Truck className="w-5 h-5" />} 
            label={t.vehicles} 
            active={activeTab === 'vehicles'} 
            onClick={() => setActiveTab('vehicles')} 
          />
          <NavItem 
            icon={<BarChart3 className="w-5 h-5" />} 
            label={t.results} 
            active={activeTab === 'results'} 
            onClick={() => setActiveTab('results')} 
            disabled={!result}
          />
        </nav>

        <div className="absolute bottom-8 left-0 w-full px-6 space-y-4">
          <div className="relative group">
            <div className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 pointer-events-none">
              <Globe className="w-4 h-4" />
            </div>
            <select 
              value={lang}
              onChange={(e) => setLang(e.target.value as Language)}
              className="w-full bg-gray-50 border border-gray-200 text-gray-600 text-sm rounded-xl focus:ring-emerald-500 focus:border-emerald-500 block p-2.5 pl-10 appearance-none cursor-pointer hover:bg-gray-100 transition-all font-medium"
            >
              <option value="zh">简体中文 (ZH)</option>
              <option value="en">English (EN)</option>
              <option value="fr">Français (FR)</option>
              <option value="de">Deutsch (DE)</option>
              <option value="it">Italiano (IT)</option>
            </select>
            <div className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 pointer-events-none">
              <ChevronDown className="w-4 h-4" />
            </div>
          </div>
          <button 
            onClick={handleOptimize}
            className="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-3 rounded-xl font-semibold flex items-center justify-center gap-2 transition-all shadow-lg shadow-emerald-600/20"
          >
            <Play className="w-4 h-4 fill-current" />
            {t.startCalc}
          </button>
        </div>
      </aside>

      {/* Mobile Bottom Nav */}
      <nav className="md:hidden fixed bottom-0 left-0 w-full bg-white border-t border-[#E5E5E5] px-2 py-2 flex justify-around items-center z-30 shadow-[0_-4px_12px_rgba(0,0,0,0.05)]">
        <MobileNavItem 
          icon={<LayoutDashboard className="w-6 h-6" />} 
          label={t.home} 
          active={activeTab === 'dashboard'} 
          onClick={() => setActiveTab('dashboard')} 
        />
        <MobileNavItem 
          icon={<Package className="w-6 h-6" />} 
          label={t.goodsTab} 
          active={activeTab === 'goods'} 
          onClick={() => setActiveTab('goods')} 
        />
        <MobileNavItem 
          icon={<Truck className="w-6 h-6" />} 
          label={t.vehiclesTab} 
          active={activeTab === 'vehicles'} 
          onClick={() => setActiveTab('vehicles')} 
        />
        <MobileNavItem 
          icon={<BarChart3 className="w-6 h-6" />} 
          label={t.resultsTab} 
          active={activeTab === 'results'} 
          onClick={() => setActiveTab('results')} 
          disabled={!result}
        />
        <button 
          onClick={handleOptimize}
          className="bg-emerald-600 text-white p-3 rounded-full shadow-lg shadow-emerald-600/30 active:scale-95 transition-transform"
        >
          <Play className="w-6 h-6 fill-current" />
        </button>
      </nav>

      {/* Mobile Header */}
      <header className="md:hidden sticky top-0 bg-white/80 backdrop-blur-md border-b border-[#E5E5E5] px-4 py-3 flex justify-between items-center z-20">
        <div className="flex items-center gap-2 text-emerald-600 font-bold">
          <Truck className="w-6 h-6" />
          <span className="text-lg">{t.appName}</span>
        </div>
        <div className="relative">
          <select 
            value={lang}
            onChange={(e) => setLang(e.target.value as Language)}
            className="bg-gray-100 border-none text-gray-600 text-[10px] font-bold rounded-lg focus:ring-0 block py-1.5 pl-2 pr-6 appearance-none cursor-pointer"
          >
            <option value="zh">ZH</option>
            <option value="en">EN</option>
            <option value="fr">FR</option>
            <option value="de">DE</option>
            <option value="it">IT</option>
          </select>
          <div className="absolute right-1.5 top-1/2 -translate-y-1/2 text-gray-400 pointer-events-none">
            <ChevronDown className="w-3 h-3" />
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="md:ml-64 p-4 md:p-8">
        <AnimatePresence mode="wait">
          {activeTab === 'dashboard' && (
            <motion.div 
              key="dashboard"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8"
            >
              <header>
                <h1 className="text-3xl font-bold tracking-tight">{t.welcome}</h1>
                <p className="text-gray-500 mt-1">{t.sysStatus}</p>
              </header>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <StatCard 
                  title={t.pendingGoods} 
                  value={goods.reduce((sum, g) => sum + g.quantity, 0)} 
                  unit={t.unitPcs} 
                  icon={<Package className="text-blue-500" />}
                />
                <StatCard 
                  title={t.availableVehicles} 
                  value={vehicles.reduce((sum, v) => sum + v.count, 0)} 
                  unit={t.unitVehicles} 
                  icon={<Truck className="text-emerald-500" />}
                />
                <StatCard 
                  title={t.totalWeight} 
                  value={(goods.reduce((sum, g) => sum + (g.weight * g.quantity), 0) / 1000).toFixed(1)} 
                  unit={t.unitTons} 
                  icon={<Info className="text-amber-500" />}
                />
              </div>

              <div className="bg-white rounded-2xl p-8 border border-[#E5E5E5] shadow-sm">
                <h2 className="text-xl font-bold mb-6">{t.guideTitle}</h2>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
                  <GuideStep 
                    number="01" 
                    title={t.guideStep1Title} 
                    desc={t.guideStep1Desc} 
                  />
                  <GuideStep 
                    number="02" 
                    title={t.guideStep2Title} 
                    desc={t.guideStep2Desc} 
                  />
                  <GuideStep 
                    number="03" 
                    title={t.guideStep3Title} 
                    desc={t.guideStep3Desc} 
                  />
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'goods' && (
            <motion.div 
              key="goods"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-6"
            >
              <div className="flex flex-col md:flex-row justify-between items-start md:items-end gap-4">
                <div>
                  <h1 className="text-2xl md:text-3xl font-bold tracking-tight">{t.goodsList}</h1>
                  <p className="text-gray-500 mt-1">{t.goodsListDesc}</p>
                </div>
                <div className="flex flex-wrap gap-2 w-full md:w-auto">
                  {selectedGoods.size > 0 && (
                    <button 
                      onClick={deleteSelectedGoods}
                      className="bg-red-50 text-red-600 border border-red-100 hover:bg-red-100 px-3 py-2 rounded-lg flex items-center gap-2 font-medium transition-colors text-sm"
                    >
                      <Trash2 className="w-4 h-4" />
                      <span className="hidden sm:inline">{t.deleteSelected}</span> ({selectedGoods.size})
                    </button>
                  )}
                  <label className="bg-white border border-[#E5E5E5] hover:bg-gray-50 px-3 py-2 rounded-lg flex items-center gap-2 font-medium transition-colors cursor-pointer text-sm">
                    <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                    {t.importExcel}
                    <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={handleExcelImport} />
                  </label>
                  <button 
                    onClick={downloadTemplate}
                    className="bg-white border border-[#E5E5E5] hover:bg-gray-50 px-3 py-2 rounded-lg flex items-center gap-2 font-medium transition-colors text-sm"
                  >
                    <Download className="w-4 h-4 text-blue-600" />
                    {t.downloadTemplate}
                  </button>
                  <button 
                    onClick={addGood}
                    className="bg-emerald-600 text-white hover:bg-emerald-700 px-3 py-2 rounded-lg flex items-center gap-2 font-medium transition-colors text-sm flex-1 md:flex-none justify-center"
                  >
                    <Plus className="w-4 h-4" />
                    {t.manualAdd}
                  </button>
                </div>
              </div>

              {/* Mobile Card List */}
              <div className="md:hidden space-y-4">
                <div className="flex items-center gap-2 px-2">
                  <input 
                    type="checkbox" 
                    checked={selectedGoods.size === goods.length && goods.length > 0}
                    onChange={toggleAllGoods}
                    className="rounded border-gray-300 text-emerald-600 focus:ring-emerald-500 w-5 h-5"
                  />
                  <span className="text-sm font-medium text-gray-500">{t.selectAll}</span>
                </div>
                {goods.map((good) => (
                  <div key={good.id} className={`bg-white p-4 rounded-xl border ${selectedGoods.has(good.id) ? 'border-emerald-500 bg-emerald-50/20' : 'border-[#E5E5E5]'} shadow-sm space-y-3`}>
                    <div className="flex justify-between items-start">
                      <div className="flex items-center gap-3">
                        <input 
                          type="checkbox" 
                          checked={selectedGoods.has(good.id)}
                          onChange={() => toggleGoodSelection(good.id)}
                          className="rounded border-gray-300 text-emerald-600 focus:ring-emerald-500 w-5 h-5"
                        />
                        <input 
                          type="text" 
                          value={good.name} 
                          onChange={(e) => updateGood(good.id, 'name', e.target.value)}
                          className="font-bold bg-transparent border-none focus:ring-0 p-0"
                        />
                      </div>
                      <button onClick={() => removeGood(good.id)} className="text-red-400 p-1"><Trash2 className="w-4 h-4" /></button>
                    </div>
                    <div className="grid grid-cols-3 gap-4 text-sm">
                      <div>
                        <p className="text-gray-400 text-[10px] uppercase font-bold">{t.weightPerUnit.split(' ')[0]}</p>
                        <input 
                          type="number" 
                          value={good.weight} 
                          onChange={(e) => updateGood(good.id, 'weight', parseFloat(e.target.value) || 0)}
                          className="w-full bg-gray-50 rounded px-2 py-1 mt-1"
                        />
                      </div>
                      <div>
                        <p className="text-gray-400 text-[10px] uppercase font-bold">{t.volumePerUnit.split(' ')[0]}</p>
                        <input 
                          type="number" 
                          step="0.01"
                          value={good.volume} 
                          onChange={(e) => updateGood(good.id, 'volume', parseFloat(e.target.value) || 0)}
                          className="w-full bg-gray-50 rounded px-2 py-1 mt-1"
                        />
                      </div>
                      <div>
                        <p className="text-gray-400 text-[10px] uppercase font-bold">{t.quantity}</p>
                        <input 
                          type="number" 
                          value={good.quantity} 
                          onChange={(e) => updateGood(good.id, 'quantity', parseInt(e.target.value) || 0)}
                          className="w-full bg-gray-50 rounded px-2 py-1 mt-1"
                        />
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              {/* Desktop Table View */}
              <div className="hidden md:block bg-white rounded-2xl border border-[#E5E5E5] overflow-hidden shadow-sm">
                <table className="w-full text-left">
                  <thead className="bg-gray-50 border-b border-[#E5E5E5]">
                    <tr>
                      <th className="px-6 py-4 w-10">
                        <input 
                          type="checkbox" 
                          checked={selectedGoods.size === goods.length && goods.length > 0}
                          onChange={toggleAllGoods}
                          className="rounded border-gray-300 text-emerald-600 focus:ring-emerald-500"
                        />
                      </th>
                      <th className="px-6 py-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">{t.goodName}</th>
                      <th className="px-6 py-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">{t.weightPerUnit}</th>
                      <th className="px-6 py-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">{t.volumePerUnit}</th>
                      <th className="px-6 py-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">{t.quantity}</th>
                      <th className="px-6 py-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">{t.actions}</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E5E5]">
                    {goods.map((good) => (
                      <tr key={good.id} className={`hover:bg-gray-50/50 transition-colors ${selectedGoods.has(good.id) ? 'bg-emerald-50/30' : ''}`}>
                        <td className="px-6 py-4">
                          <input 
                            type="checkbox" 
                            checked={selectedGoods.has(good.id)}
                            onChange={() => toggleGoodSelection(good.id)}
                            className="rounded border-gray-300 text-emerald-600 focus:ring-emerald-500"
                          />
                        </td>
                        <td className="px-6 py-4">
                          <input 
                            type="text" 
                            value={good.name} 
                            onChange={(e) => updateGood(good.id, 'name', e.target.value)}
                            className="w-full bg-transparent border-none focus:ring-0 font-medium"
                          />
                        </td>
                        <td className="px-6 py-4">
                          <input 
                            type="number" 
                            value={good.weight} 
                            onChange={(e) => updateGood(good.id, 'weight', parseFloat(e.target.value) || 0)}
                            className="w-24 bg-transparent border-none focus:ring-0"
                          />
                        </td>
                        <td className="px-6 py-4">
                          <input 
                            type="number" 
                            step="0.01"
                            value={good.volume} 
                            onChange={(e) => updateGood(good.id, 'volume', parseFloat(e.target.value) || 0)}
                            className="w-24 bg-transparent border-none focus:ring-0"
                          />
                        </td>
                        <td className="px-6 py-4">
                          <input 
                            type="number" 
                            value={good.quantity} 
                            onChange={(e) => updateGood(good.id, 'quantity', parseInt(e.target.value) || 0)}
                            className="w-24 bg-transparent border-none focus:ring-0"
                          />
                        </td>
                        <td className="px-6 py-4">
                          <button 
                            onClick={() => removeGood(good.id)}
                            className="text-red-400 hover:text-red-600 p-2 rounded-lg hover:bg-red-50 transition-all"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </motion.div>
          )}

          {activeTab === 'vehicles' && (
            <motion.div 
              key="vehicles"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-6"
            >
              <div className="flex flex-col md:flex-row justify-between items-start md:items-end gap-4">
                <div>
                  <h1 className="text-2xl md:text-3xl font-bold tracking-tight">{t.vehicleConfig}</h1>
                  <p className="text-gray-500 mt-1">{t.vehicleConfigDesc}</p>
                </div>
                <div className="flex gap-2 w-full md:w-auto">
                  {selectedVehicles.size > 0 && (
                    <button 
                      onClick={deleteSelectedVehicles}
                      className="bg-red-50 text-red-600 border border-red-100 hover:bg-red-100 px-3 py-2 rounded-lg flex items-center gap-2 font-medium transition-colors text-sm"
                    >
                      <Trash2 className="w-4 h-4" />
                      <span className="hidden sm:inline">{t.deleteSelected}</span> ({selectedVehicles.size})
                    </button>
                  )}
                  <button 
                    onClick={addVehicle}
                    className="bg-emerald-600 text-white hover:bg-emerald-700 px-3 py-2 rounded-lg flex items-center gap-2 font-medium transition-colors text-sm flex-1 md:flex-none justify-center"
                  >
                    <Plus className="w-4 h-4" />
                    {t.addVehicleType}
                  </button>
                </div>
              </div>

              {/* Mobile Card List */}
              <div className="md:hidden space-y-4">
                <div className="flex items-center gap-2 px-2">
                  <input 
                    type="checkbox" 
                    checked={selectedVehicles.size === vehicles.length && vehicles.length > 0}
                    onChange={toggleAllVehicles}
                    className="rounded border-gray-300 text-emerald-600 focus:ring-emerald-500 w-5 h-5"
                  />
                  <span className="text-sm font-medium text-gray-500">{t.selectAll}</span>
                </div>
                {vehicles.map((vehicle) => (
                  <div key={vehicle.id} className={`bg-white p-4 rounded-xl border ${selectedVehicles.has(vehicle.id) ? 'border-emerald-500 bg-emerald-50/20' : 'border-[#E5E5E5]'} shadow-sm space-y-3`}>
                    <div className="flex justify-between items-start">
                      <div className="flex items-center gap-3">
                        <input 
                          type="checkbox" 
                          checked={selectedVehicles.has(vehicle.id)}
                          onChange={() => toggleVehicleSelection(vehicle.id)}
                          className="rounded border-gray-300 text-emerald-600 focus:ring-emerald-500 w-5 h-5"
                        />
                        <input 
                          type="text" 
                          value={vehicle.type} 
                          onChange={(e) => updateVehicle(vehicle.id, 'type', e.target.value)}
                          className="font-bold bg-transparent border-none focus:ring-0 p-0"
                        />
                      </div>
                      <button onClick={() => removeVehicle(vehicle.id)} className="text-red-400 p-1"><Trash2 className="w-4 h-4" /></button>
                    </div>
                    <div className="grid grid-cols-3 gap-4 text-sm">
                      <div>
                        <p className="text-gray-400 text-[10px] uppercase font-bold">载重 (kg)</p>
                        <input 
                          type="number" 
                          value={vehicle.maxWeight} 
                          onChange={(e) => updateVehicle(vehicle.id, 'maxWeight', parseFloat(e.target.value) || 0)}
                          className="w-full bg-gray-50 rounded px-2 py-1 mt-1"
                        />
                      </div>
                      <div>
                        <p className="text-gray-400 text-[10px] uppercase font-bold">容积 (m³)</p>
                        <input 
                          type="number" 
                          value={vehicle.maxVolume} 
                          onChange={(e) => updateVehicle(vehicle.id, 'maxVolume', parseFloat(e.target.value) || 0)}
                          className="w-full bg-gray-50 rounded px-2 py-1 mt-1"
                        />
                      </div>
                      <div>
                        <p className="text-gray-400 text-[10px] uppercase font-bold">数量</p>
                        <input 
                          type="number" 
                          value={vehicle.count} 
                          onChange={(e) => updateVehicle(vehicle.id, 'count', parseInt(e.target.value) || 0)}
                          className="w-full bg-gray-50 rounded px-2 py-1 mt-1"
                        />
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              {/* Desktop Table View */}
              <div className="hidden md:block bg-white rounded-2xl border border-[#E5E5E5] overflow-hidden shadow-sm">
                <table className="w-full text-left">
                  <thead className="bg-gray-50 border-b border-[#E5E5E5]">
                    <tr>
                      <th className="px-6 py-4 w-10">
                        <input 
                          type="checkbox" 
                          checked={selectedVehicles.size === vehicles.length && vehicles.length > 0}
                          onChange={toggleAllVehicles}
                          className="rounded border-gray-300 text-emerald-600 focus:ring-emerald-500"
                        />
                      </th>
                      <th className="px-6 py-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">{t.vehicleTypeName}</th>
                      <th className="px-6 py-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">{t.maxWeight}</th>
                      <th className="px-6 py-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">{t.maxVolume}</th>
                      <th className="px-6 py-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">{t.vehicleCount}</th>
                      <th className="px-6 py-4 text-xs font-semibold text-gray-400 uppercase tracking-wider">{t.actions}</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E5E5]">
                    {vehicles.map((vehicle) => (
                      <tr key={vehicle.id} className={`hover:bg-gray-50/50 transition-colors ${selectedVehicles.has(vehicle.id) ? 'bg-emerald-50/30' : ''}`}>
                        <td className="px-6 py-4">
                          <input 
                            type="checkbox" 
                            checked={selectedVehicles.has(vehicle.id)}
                            onChange={() => toggleVehicleSelection(vehicle.id)}
                            className="rounded border-gray-300 text-emerald-600 focus:ring-emerald-500"
                          />
                        </td>
                        <td className="px-6 py-4">
                          <input 
                            type="text" 
                            value={vehicle.type} 
                            onChange={(e) => updateVehicle(vehicle.id, 'type', e.target.value)}
                            className="w-full bg-transparent border-none focus:ring-0 font-medium"
                          />
                        </td>
                        <td className="px-6 py-4">
                          <input 
                            type="number" 
                            value={vehicle.maxWeight} 
                            onChange={(e) => updateVehicle(vehicle.id, 'maxWeight', parseFloat(e.target.value) || 0)}
                            className="w-32 bg-transparent border-none focus:ring-0"
                          />
                        </td>
                        <td className="px-6 py-4">
                          <input 
                            type="number" 
                            value={vehicle.maxVolume} 
                            onChange={(e) => updateVehicle(vehicle.id, 'maxVolume', parseFloat(e.target.value) || 0)}
                            className="w-32 bg-transparent border-none focus:ring-0"
                          />
                        </td>
                        <td className="px-6 py-4">
                          <input 
                            type="number" 
                            value={vehicle.count} 
                            onChange={(e) => updateVehicle(vehicle.id, 'count', parseInt(e.target.value) || 0)}
                            className="w-24 bg-transparent border-none focus:ring-0"
                          />
                        </td>
                        <td className="px-6 py-4">
                          <button 
                            onClick={() => removeVehicle(vehicle.id)}
                            className="text-red-400 hover:text-red-600 p-2 rounded-lg hover:bg-red-50 transition-all"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </motion.div>
          )}

          {activeTab === 'results' && result && (
            <motion.div 
              key="results"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-8"
            >
              <header className="flex justify-between items-end">
                <div>
                  <h1 className="text-3xl font-bold tracking-tight">{t.optimizationPlan}</h1>
                  <p className="text-gray-500 mt-1">{t.optimizationPlanDesc}</p>
                </div>
                <div className="flex gap-4">
                  <div className="text-right">
                    <p className="text-xs text-gray-400 uppercase font-bold">{t.avgWeightUtil}</p>
                    <p className="text-xl font-bold text-emerald-600">{result.summary.avgWeightUtilization.toFixed(1)}%</p>
                  </div>
                  <div className="text-right">
                    <p className="text-xs text-gray-400 uppercase font-bold">{t.avgVolumeUtil}</p>
                    <p className="text-xl font-bold text-blue-600">{result.summary.avgVolumeUtilization.toFixed(1)}%</p>
                  </div>
                </div>
              </header>

              {result.unassignedGoods.length > 0 && (
                <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex items-start gap-3">
                  <AlertCircle className="w-5 h-5 text-amber-500 mt-0.5" />
                  <div>
                    <h4 className="font-bold text-amber-800">{t.unassignedAlert}</h4>
                    <p className="text-sm text-amber-700">{t.unassignedDesc} {result.unassignedGoods.map(g => `${g.name} x${g.quantity}`).join(', ')}</p>
                  </div>
                </div>
              )}

              <div className="grid grid-cols-1 gap-6">
                {result.loads.map((load, idx) => (
                  <div key={load.vehicleId} className="bg-white rounded-2xl border border-[#E5E5E5] overflow-hidden shadow-sm">
                    <div className="p-6 border-b border-[#E5E5E5] flex justify-between items-center bg-gray-50/50">
                      <div className="flex items-center gap-4">
                        <div className="w-12 h-12 bg-emerald-100 rounded-xl flex items-center justify-center text-emerald-600">
                          <Truck className="w-6 h-6" />
                        </div>
                        <div>
                          <h3 className="font-bold text-lg">{t.vehicleLabel} #{idx + 1}: {load.vehicleType}</h3>
                          <p className="text-sm text-gray-500">ID: {load.vehicleId}</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-8">
                        <UtilizationBar label={t.weightLoad} value={load.weightUtilization} color="bg-emerald-500" />
                        <UtilizationBar label={t.volumeLoad} value={load.volumeUtilization} color="bg-blue-500" />
                        <button 
                          onClick={() => deleteLoad(load.vehicleId)}
                          className="text-red-400 hover:text-red-600 p-2 rounded-lg hover:bg-red-50 transition-all ml-4"
                          title={t.deleteLoad}
                        >
                          <Trash2 className="w-5 h-5" />
                        </button>
                      </div>
                    </div>
                    <div className="p-6">
                      <table className="w-full text-left text-sm">
                        <thead>
                          <tr className="text-gray-400 font-medium border-b border-[#F0F0F0]">
                            <th className="pb-3">{t.goodName}</th>
                            <th className="pb-3">{t.quantity}</th>
                            <th className="pb-3">{t.totalWeight}</th>
                            <th className="pb-3">{t.totalVolume}</th>
                            <th className="pb-3 text-right">{t.actions}</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-[#F0F0F0]">
                          {load.items.map((item, i) => (
                            <tr key={i}>
                              <td className="py-3 font-medium">{item.goodName}</td>
                              <td className="py-3">
                                <input 
                                  type="number" 
                                  value={item.quantity} 
                                  onChange={(e) => updateLoadItem(load.vehicleId, item.goodId, parseInt(e.target.value) || 0)}
                                  className="w-20 bg-gray-50 border border-gray-200 rounded px-2 py-1 focus:ring-1 focus:ring-emerald-500"
                                />
                                <span className="ml-1 text-gray-400">{t.unitPcs}</span>
                              </td>
                              <td className="py-3">{item.weight.toFixed(1)} kg</td>
                              <td className="py-3">{item.volume.toFixed(2)} m³</td>
                              <td className="py-3 text-right">
                                <button 
                                  onClick={() => updateLoadItem(load.vehicleId, item.goodId, 0)}
                                  className="text-gray-400 hover:text-red-500 transition-colors"
                                  title={t.removeItem}
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr className="font-bold text-gray-900 border-t border-[#E5E5E5]">
                            <td className="pt-4">{t.total}</td>
                            <td className="pt-4">{load.items.reduce((s, i) => s + i.quantity, 0)} {t.unitPcs}</td>
                            <td className="pt-4">{load.totalWeight.toFixed(1)} kg</td>
                            <td className="pt-4">{load.totalVolume.toFixed(2)} m³</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                ))}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}

function MobileNavItem({ icon, label, active, onClick, disabled = false }: { 
  icon: React.ReactNode, 
  label: string, 
  active: boolean, 
  onClick: () => void,
  disabled?: boolean
}) {
  return (
    <button 
      onClick={onClick}
      disabled={disabled}
      className={`flex flex-col items-center gap-1 transition-all ${
        disabled ? 'opacity-20 grayscale cursor-not-allowed' :
        active ? 'text-emerald-600' : 'text-gray-400'
      }`}
    >
      {icon}
      <span className="text-[10px] font-bold uppercase tracking-tighter">{label}</span>
    </button>
  );
}

function NavItem({ icon, label, active, onClick, disabled = false }: { 
  icon: React.ReactNode, 
  label: string, 
  active: boolean, 
  onClick: () => void,
  disabled?: boolean
}) {
  return (
    <button 
      onClick={onClick}
      disabled={disabled}
      className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${
        disabled ? 'opacity-30 cursor-not-allowed' :
        active ? 'bg-emerald-50 text-emerald-600 font-bold' : 'text-gray-500 hover:bg-gray-50'
      }`}
    >
      {icon}
      <span>{label}</span>
      {active && <ChevronRight className="w-4 h-4 ml-auto" />}
    </button>
  );
}

function StatCard({ title, value, unit, icon }: { title: string, value: string | number, unit: string, icon: React.ReactNode }) {
  return (
    <div className="bg-white p-6 rounded-2xl border border-[#E5E5E5] shadow-sm">
      <div className="flex justify-between items-start">
        <div>
          <p className="text-xs text-gray-400 uppercase font-bold tracking-wider">{title}</p>
          <div className="mt-2 flex items-baseline gap-1">
            <span className="text-3xl font-bold">{value}</span>
            <span className="text-gray-400 text-sm">{unit}</span>
          </div>
        </div>
        <div className="p-3 bg-gray-50 rounded-xl">
          {icon}
        </div>
      </div>
    </div>
  );
}

function GuideStep({ number, title, desc }: { number: string, title: string, desc: string }) {
  return (
    <div className="space-y-3">
      <div className="text-4xl font-black text-gray-100 font-mono">{number}</div>
      <h3 className="font-bold text-lg">{title}</h3>
      <p className="text-sm text-gray-500 leading-relaxed">{desc}</p>
    </div>
  );
}

function UtilizationBar({ label, value, color }: { label: string, value: number, color: string }) {
  return (
    <div className="w-32">
      <div className="flex justify-between text-[10px] font-bold uppercase text-gray-400 mb-1">
        <span>{label}</span>
        <span>{value.toFixed(0)}%</span>
      </div>
      <div className="h-2 bg-gray-100 rounded-full overflow-hidden">
        <motion.div 
          initial={{ width: 0 }}
          animate={{ width: `${Math.min(value, 100)}%` }}
          className={`h-full ${color}`}
        />
      </div>
    </div>
  );
}
