# -*- coding: utf-8 -*-
"""
Objective 4: Calculate Cost-HCL Reduction Effectiveness for each loop

Simple script to:
1. Read HCL reduction data from "HCL Reduction.xlsx"
2. Read cost data from existing PVI analysis
3. Calculate Cost-HCL Reduction Effectiveness (Total only)
4. Create visualization
"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

# File paths
BASE_FOLDER = r"D:\Ph. D\Task-1\Data-Ahvaz"
TEST_DATASET_FOLDER = os.path.join(BASE_FOLDER, "Test Dataset")

# Input files
SWMM_ANALYSIS_EXCEL = os.path.join(TEST_DATASET_FOLDER, "SWMM_Enhanced_Analysis_Top30.xlsx")
HCL_REDUCTION_EXCEL = os.path.join(TEST_DATASET_FOLDER, "HCL Reduction.xlsx")

def load_original_hcl_data():
    """Load original HCL data to get baseline values and full HCL distribution"""
    try:
        df_all = pd.read_excel(SWMM_ANALYSIS_EXCEL, sheet_name='All_Results')
        original_hcl_avg = df_all['HCL'].mean()
        original_hcl_max = df_all['HCL'].max()
        original_hcl_total = df_all['HCL'].sum()
        original_hcl_values = df_all['HCL'].values
        print(f"Original HCL - Average: {original_hcl_avg:.4f}, Maximum: {original_hcl_max:.4f}")
        print(f"Original HCL - TOTAL: {original_hcl_total:.4f}")
        print(f"Total pipes with HCL data: {len(original_hcl_values)}")
        return original_hcl_avg, original_hcl_max, original_hcl_total, original_hcl_values
    except Exception as e:
        print(f"Error loading original HCL data: {e}")
        synthetic_hcl = np.random.uniform(2.0, 6.0, 100)
        return 5.5, 6.0, 550.0, synthetic_hcl

def load_hcl_reduction_data():
    """Load HCL reduction data from Excel file"""
    try:
        df_hcl = pd.read_excel(HCL_REDUCTION_EXCEL)
        print(f"Loaded {len(df_hcl)} loops of HCL reduction data")
        print("Columns:", df_hcl.columns.tolist())
        return df_hcl
    except Exception as e:
        print(f"Error loading HCL reduction data: {e}")
        return pd.DataFrame({
            'Loop': range(1, 51),
            'Total value reduction ratio': np.random.uniform(0.01, 0.15, 50),
            'Maximum value reduction ratio': np.random.uniform(0.02, 0.25, 50)
        })

def get_cost_data():
    """Read actual cost data from PVI analysis results"""
    try:
        pvi_results_path = os.path.join(TEST_DATASET_FOLDER, 'FIXED_PVI_Cost_Effectiveness_Analysis.xlsx')
        df_pvi_cycles = pd.read_excel(pvi_results_path, sheet_name='FIXED_Cycle_Results')
        costs = df_pvi_cycles['Cost'].tolist()
        
        print(f"Loaded {len(costs)} actual cost values from PVI analysis")
        print(f"Cost range: ${min(costs):,.2f} to ${max(costs):,.2f}")
        print(f"First 5 costs: {[f'${c:,.2f}' for c in costs[:5]]}")
        
        return costs
        
    except Exception as e:
        print(f"Error loading cost data: {e}")
        print("Using fallback synthetic cost data")
        
        base_cost = 30000
        costs = []
        for i in range(50):
            cost = base_cost + (i * 2000) + np.random.normal(0, 3000)
            cost = max(cost, 15000)
            costs.append(cost)
        return costs

def calculate_cost_hcl_effectiveness(hcl_reduction_df, costs, original_hcl_total):
    """
    Calculate Cost-HCL Reduction Effectiveness for each loop
    
    IMPORTANT CLARIFICATIONS:
    - Reduction ratios are CUMULATIVE (relative to original baseline)
    - Example: 0.03 means 3% reduction from original, 0.21 means 21% reduction from original
    - We calculate INCREMENTAL reduction for each cycle (difference between consecutive cycles)
    - Cost is per cycle (not cumulative)
    - Formula: Cost-HCL Effectiveness = (HCL_reduction_this_cycle / Cost_this_cycle) × 10,000
    """
    
    results = []
    
    print("\n" + "="*60)
    print("COST-HCL EFFECTIVENESS CALCULATION")
    print("="*60)
    print("Formula: Cost-HCL Effectiveness = (HCL_reduction_this_cycle) / (Cost_this_cycle) × 10,000")
    print("NOTE: HCL reduction is INCREMENTAL (cycle-to-cycle), Cost is per cycle")
    print(f"Original Total HCL: {original_hcl_total:.4f}")
    print("="*60)
    
    # Track previous cycle value for calculating incremental changes
    previous_total_hcl = original_hcl_total
    
    for idx, row in hcl_reduction_df.iterrows():
        loop = int(row['Loop'])
        
        # This is CUMULATIVE reduction ratio (relative to original)
        total_reduction_ratio = row['Total value reduction ratio']
        
        # Get cost for this specific cycle
        if loop <= len(costs):
            cycle_cost = costs[loop - 1]
        else:
            print(f"Warning: No cost data for loop {loop}")
            continue
        
        # Calculate CURRENT state (after this cycle's reduction)
        # Relative to ORIGINAL baseline
        current_total_hcl = original_hcl_total * (1 - total_reduction_ratio)
        
        # Calculate INCREMENTAL HCL reduction for THIS CYCLE ONLY
        cycle_hcl_reduction = previous_total_hcl - current_total_hcl
        
        # Calculate Cost-HCL Effectiveness for THIS CYCLE
        if cycle_cost > 0:
            cost_hcl_effectiveness = (cycle_hcl_reduction / cycle_cost) * 10000
        else:
            cost_hcl_effectiveness = 0
        
        # Debug output for first few cycles
        if loop <= 5:
            print(f"\nLoop {loop} Details:")
            print(f"  Cumulative reduction ratio: {total_reduction_ratio:.4f} ({total_reduction_ratio*100:.2f}%)")
            print(f"  Previous total HCL: {previous_total_hcl:.4f}")
            print(f"  Current total HCL: {current_total_hcl:.4f}")
            print(f"  Incremental HCL reduction: {cycle_hcl_reduction:.4f}")
            print(f"  Cost THIS CYCLE: ${cycle_cost:,.2f}")
            print(f"  Effectiveness: ({cycle_hcl_reduction:.4f} / {cycle_cost:.2f}) × 10,000 = {cost_hcl_effectiveness:.4f}")
        
        results.append({
            'Loop': int(loop),
            'Cumulative_Reduction_Ratio': total_reduction_ratio,
            'Incremental_HCL_Reduction': cycle_hcl_reduction,
            'Current_Total_HCL': current_total_hcl,
            'Cycle_Cost': cycle_cost,
            'Cost_HCL_Effectiveness': cost_hcl_effectiveness
        })
        
        # Update for next iteration
        previous_total_hcl = current_total_hcl
    
    results_df = pd.DataFrame(results)
    
    # Print summary
    print("\n" + "="*60)
    print("CALCULATION SUMMARY")
    print("="*60)
    print(f"Total loops processed: {len(results_df)}")
    print(f"Total incremental HCL reduction: {results_df['Incremental_HCL_Reduction'].sum():.4f}")
    print(f"Total cost: ${sum(costs[:len(results_df)]):,.2f}")
    print(f"Average effectiveness: {results_df['Cost_HCL_Effectiveness'].mean():.4f}")
    print("="*60)
    
    return results_df

def load_hcl_reduction_data():
    """Load HCL reduction data from Excel file"""
    try:
        print(f"\nAttempting to read: {HCL_REDUCTION_EXCEL}")
        print(f"File exists: {os.path.exists(HCL_REDUCTION_EXCEL)}")
        
        df_hcl = pd.read_excel(HCL_REDUCTION_EXCEL)
        
        print(f"Successfully loaded {len(df_hcl)} loops of HCL reduction data")
        print("Columns found:", df_hcl.columns.tolist())
        print("\nFirst 10 rows of data:")
        print(df_hcl.head(10))
        
        # Verify the data matches what you expect
        if 'Total value reduction ratio' in df_hcl.columns:
            print(f"\nFirst 5 'Total value reduction ratio' values:")
            print(df_hcl['Total value reduction ratio'].head().tolist())
        
        return df_hcl
        
    except Exception as e:
        print(f"ERROR loading HCL reduction data: {e}")
        print("Creating SYNTHETIC data - THIS IS NOT YOUR REAL DATA!")
        return pd.DataFrame({
            'Loop': range(1, 51),
            'Total value reduction ratio': np.random.uniform(0.01, 0.15, 50),
            'Maximum value reduction ratio': np.random.uniform(0.02, 0.25, 50)
        })

def plot_cost_hcl_effectiveness(results_df, original_hcl_values, original_hcl_max, hcl_reduction_df, save_path=None):
    """Create plot showing HCL reduction effectiveness across cycles"""
    
    # Use ALL available loops (up to 50)
    n_loops = min(50, len(results_df), len(hcl_reduction_df))
    plot_df = results_df.head(n_loops)
    hcl_df = hcl_reduction_df.head(n_loops)
    
    # Create adjusted HCL data for boxplot based on CUMULATIVE reduction ratios
    boxplot_data = {}
    
    # Add original network data as first column
    boxplot_data["Original"] = original_hcl_values
    
    # Create adjusted HCL values for ALL loops using CUMULATIVE reductions
    for i in range(n_loops):
        loop_num = i + 1
        
        if i < len(hcl_df):
            # Get CUMULATIVE reduction ratio (relative to original)
            total_reduction_ratio = hcl_df.iloc[i]['Total value reduction ratio']
            
            # Apply cumulative reduction to original data
            adjusted_hcl = original_hcl_values * (1 - total_reduction_ratio)
            
            boxplot_data[f"Loop {loop_num}"] = adjusted_hcl
    
    # Convert to DataFrame for boxplot
    df_boxplot = pd.DataFrame(boxplot_data)
    
    # Create figure
    plt.figure(dpi=300, figsize=(15, 6))
    
    # Draw boxplot
    boxplot = df_boxplot.boxplot(grid=False)
    plt.xticks(rotation=90, ha='center', fontsize=8)
    plt.ylabel('HCL Values', fontsize=12)
    
    # Get current axis and create second y-axis for effectiveness values
    ax1 = plt.gca()
    ax2 = ax1.twinx()
    
    # x-axis positions - align with boxplot positions (1-based indexing)
    n_boxes = len(df_boxplot.columns)
    x_positions = range(1, n_boxes + 1)
    
    # Create effectiveness values - original (0) + effectiveness for each loop
    efficiency_values = [0]  # Original network has 0 effectiveness (no cost)
    
    # Add effectiveness values for loops
    for i in range(min(len(plot_df), n_boxes - 1)):
        eff_value = plot_df.iloc[i]['Cost_HCL_Effectiveness']
        efficiency_values.append(eff_value)
    
    # Ensure we have the right number of efficiency values
    while len(efficiency_values) < n_boxes:
        efficiency_values.append(0)
    
    # Plot scatter on second y-axis showing effectiveness values
    ax2.scatter(x_positions, efficiency_values[:n_boxes], color='red', 
                label='Unit cost-effectiveness (HCL/10k USD)', zorder=3, s=80)
    ax2.set_ylabel('Unit cost-effectiveness (HCL/10k USD)', color='red')
    ax2.tick_params(axis="y", labelcolor='red')
    ax2.legend(loc='upper right', fontsize='small')
    
    # Set left y-axis range to 0-7
    ax1.set_ylim(0, 7)
    
    # Adjust right y-axis range based on data
    max_eff = max(efficiency_values)
    ax2.set_ylim(0, 10)
    
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"Plot saved to: {save_path}")
    
    plt.show()

def main():
    """Main function to run the analysis"""
    print("="*60)
    print("OBJECTIVE 4: Cost-HCL Reduction Effectiveness Analysis")
    print("="*60)
    
    # Step 1: Load original HCL data
    print("Step 1: Loading original HCL data...")
    original_hcl_avg, original_hcl_max, original_hcl_total, original_hcl_values = load_original_hcl_data()
    
    # Step 2: Load HCL reduction data
    print("Step 2: Loading HCL reduction data...")
    hcl_reduction_df = load_hcl_reduction_data()
    
    # Step 3: Get cost data
    print("Step 3: Getting cost data...")
    costs = get_cost_data()
    
    # Step 4: Calculate Cost-HCL Effectiveness
    print("Step 4: Calculating Cost-HCL Effectiveness...")
    results_df = calculate_cost_hcl_effectiveness(hcl_reduction_df, costs, original_hcl_total)
    
    # Step 5: Display results
    print("\nStep 5: Results summary (first 10 loops):")
    print(results_df.head(10)[['Loop', 'Cost_HCL_Effectiveness', 'Incremental_HCL_Reduction', 'Cycle_Cost']].to_string(index=False))
    
    # Step 6: Create plot
    print("\nStep 6: Creating plot...")
    save_path = os.path.join(TEST_DATASET_FOLDER, 'Cost_HCL_Effectiveness_Plot.png')
    plot_cost_hcl_effectiveness(results_df, original_hcl_values, original_hcl_max, hcl_reduction_df, save_path)
    
    # Step 7: Save results to Excel
    print("\nStep 7: Saving results...")
    excel_path = os.path.join(TEST_DATASET_FOLDER, 'Cost_HCL_Effectiveness_Results.xlsx')
    results_df.to_excel(excel_path, index=False)
    print(f"Results saved to: {excel_path}")
    
    print("\nObjective 4 completed successfully!")

if __name__ == "__main__":
    main()