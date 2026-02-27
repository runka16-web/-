export interface Good {
  id: string;
  name: string;
  weight: number; // in kg
  volume: number; // in m³
  quantity: number;
}

export interface Vehicle {
  id: string;
  type: string;
  maxWeight: number; // in kg
  maxVolume: number; // in m³
  count: number;
}

export interface LoadingItem {
  goodId: string;
  goodName: string;
  quantity: number;
  weight: number;
  volume: number;
}

export interface VehicleLoad {
  vehicleId: string;
  vehicleType: string;
  items: LoadingItem[];
  totalWeight: number;
  totalVolume: number;
  weightUtilization: number; // percentage
  volumeUtilization: number; // percentage
}

export interface OptimizationResult {
  loads: VehicleLoad[];
  unassignedGoods: Good[];
  summary: {
    totalVehiclesUsed: number;
    totalWeight: number;
    totalVolume: number;
    avgWeightUtilization: number;
    avgVolumeUtilization: number;
  };
}
