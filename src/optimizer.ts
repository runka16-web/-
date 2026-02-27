import { Good, Vehicle, OptimizationResult, VehicleLoad, LoadingItem } from './types';

/**
 * A simple Bin Packing algorithm (First Fit Decreasing)
 * to assign goods to vehicles based on weight and volume constraints.
 */
export function optimizeLoading(goods: Good[], vehicles: Vehicle[]): OptimizationResult {
  // 1. Flatten goods into individual units (or handle by quantity)
  // For simplicity and performance, we'll handle them as groups but split if necessary
  let remainingGoods = goods.map(g => ({ ...g }));
  
  // 2. Sort vehicles by capacity (descending) to fill larger trucks first
  // Or sort by efficiency? Let's sort by maxWeight * maxVolume
  const availableVehicles = vehicles.flatMap(v => 
    Array.from({ length: v.count }, (_, i) => ({
      ...v,
      instanceId: `${v.id}-${i}`
    }))
  ).sort((a, b) => (b.maxWeight * b.maxVolume) - (a.maxWeight * a.maxVolume));

  const loads: VehicleLoad[] = [];
  
  for (const vehicle of availableVehicles) {
    const currentLoad: VehicleLoad = {
      vehicleId: vehicle.instanceId,
      vehicleType: vehicle.type,
      items: [],
      totalWeight: 0,
      totalVolume: 0,
      weightUtilization: 0,
      volumeUtilization: 0
    };

    for (let i = 0; i < remainingGoods.length; i++) {
      const good = remainingGoods[i];
      if (good.quantity <= 0) continue;

      // Calculate how many of this good can fit
      const weightLimit = vehicle.maxWeight - currentLoad.totalWeight;
      const volumeLimit = vehicle.maxVolume - currentLoad.totalVolume;

      const canFitWeight = Math.floor(weightLimit / good.weight);
      const canFitVolume = Math.floor(volumeLimit / good.volume);
      
      const fitQuantity = Math.min(good.quantity, canFitWeight, canFitVolume);

      if (fitQuantity > 0) {
        const itemWeight = fitQuantity * good.weight;
        const itemVolume = fitQuantity * good.volume;

        currentLoad.items.push({
          goodId: good.id,
          goodName: good.name,
          quantity: fitQuantity,
          weight: itemWeight,
          volume: itemVolume
        });

        currentLoad.totalWeight += itemWeight;
        currentLoad.totalVolume += itemVolume;
        good.quantity -= fitQuantity;
      }
    }

    if (currentLoad.items.length > 0) {
      currentLoad.weightUtilization = (currentLoad.totalWeight / vehicle.maxWeight) * 100;
      currentLoad.volumeUtilization = (currentLoad.totalVolume / vehicle.maxVolume) * 100;
      loads.push(currentLoad);
    }

    // Remove goods that are fully assigned
    remainingGoods = remainingGoods.filter(g => g.quantity > 0);
    if (remainingGoods.length === 0) break;
  }

  const totalWeight = loads.reduce((sum, l) => sum + l.totalWeight, 0);
  const totalVolume = loads.reduce((sum, l) => sum + l.totalVolume, 0);

  return {
    loads,
    unassignedGoods: remainingGoods,
    summary: {
      totalVehiclesUsed: loads.length,
      totalWeight,
      totalVolume,
      avgWeightUtilization: loads.length > 0 ? loads.reduce((sum, l) => sum + l.weightUtilization, 0) / loads.length : 0,
      avgVolumeUtilization: loads.length > 0 ? loads.reduce((sum, l) => sum + l.volumeUtilization, 0) / loads.length : 0
    }
  };
}
