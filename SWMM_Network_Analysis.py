# -*- coding: utf-8 -*-
"""
PART 1: SWMM Network Analysis and CPE2 Calculation
This module decomposes the SWMM model and calculates topology-based CPE2

CPE2 Formula: CPE2 = upstream_capacity / downstream_capacity
"""

import pandas as pd
import numpy as np
import os
import re
from collections import defaultdict

# File paths
BASE_FOLDER = r"D:\Ph. D\Task-1\Data-Ahvaz"
TEST_DATASET_FOLDER = os.path.join(BASE_FOLDER, "Test Dataset")

SWMM_FILE = os.path.join(BASE_FOLDER, "branched.inp")
CAPACITY_EXCEL = os.path.join(BASE_FOLDER, "Capacity_O.xlsx")
OUTPUT_FILE = os.path.join(TEST_DATASET_FOLDER, "Network_Topology_CPE2.xlsx")

# Create output folder if needed
os.makedirs(TEST_DATASET_FOLDER, exist_ok=True)


class SWMMParser:
    """Parse SWMM input files and extract network topology"""
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.junctions = {}
        self.outfalls = {}
        self.conduits = {}
        self.coordinates = {}
        
    def parse_file(self):
        """Parse the SWMM input file"""
        print(f"Parsing SWMM file: {self.file_path}")
        
        with open(self.file_path, 'r') as file:
            content = file.read()
        
        self._parse_junctions(content)
        self._parse_outfalls(content)
        self._parse_conduits(content)
        self._parse_coordinates(content)
        
        print(f"Parsed {len(self.conduits)} conduits, {len(self.junctions)} junctions, {len(self.outfalls)} outfalls")
    
    def _parse_section(self, content, section_name):
        """Extract a section from SWMM file"""
        pattern = rf'\[{section_name}\](.*?)(?=\n\[|\Z)'
        match = re.search(pattern, content, re.DOTALL | re.IGNORECASE)
        return match.group(1).strip() if match else ""
    
    def _parse_junctions(self, content):
        """Parse JUNCTIONS section"""
        section = self._parse_section(content, 'JUNCTIONS')
        for line in section.split('\n'):
            line = line.strip()
            if line and not line.startswith(';;') and not line.startswith(';'):
                parts = line.split()
                if len(parts) >= 2:
                    name = parts[0]
                    elevation = float(parts[1])
                    self.junctions[name] = {
                        'elevation': elevation,
                        'max_depth': float(parts[2]) if len(parts) > 2 else 0,
                    }
    
    def _parse_outfalls(self, content):
        """Parse OUTFALLS section"""
        section = self._parse_section(content, 'OUTFALLS')
        for line in section.split('\n'):
            line = line.strip()
            if line and not line.startswith(';;') and not line.startswith(';'):
                parts = line.split()
                if len(parts) >= 2:
                    name = parts[0]
                    elevation = float(parts[1])
                    self.outfalls[name] = {'elevation': elevation}
    
    def _parse_conduits(self, content):
        """Parse CONDUITS section"""
        section = self._parse_section(content, 'CONDUITS')
        for line in section.split('\n'):
            line = line.strip()
            if line and not line.startswith(';;') and not line.startswith(';'):
                parts = line.split()
                if len(parts) >= 6:
                    name = parts[0]
                    from_node = parts[1]
                    to_node = parts[2]
                    length = float(parts[3])
                    roughness = float(parts[4])
                    in_offset = float(parts[5])
                    out_offset = float(parts[6]) if len(parts) > 6 else 0
                    
                    self.conduits[name] = {
                        'from_node': from_node,
                        'to_node': to_node,
                        'length': length,
                        'roughness': roughness,
                        'in_offset': in_offset,
                        'out_offset': out_offset,
                    }
    
    def _parse_coordinates(self, content):
        """Parse COORDINATES section"""
        section = self._parse_section(content, 'COORDINATES')
        for line in section.split('\n'):
            line = line.strip()
            if line and not line.startswith(';;') and not line.startswith(';'):
                parts = line.split()
                if len(parts) >= 3:
                    node = parts[0]
                    x_coord = float(parts[1])
                    y_coord = float(parts[2])
                    self.coordinates[node] = {'x': x_coord, 'y': y_coord}
    
    def calculate_slopes(self):
        """Calculate slope for each conduit"""
        print("Calculating pipe slopes...")
        slopes = {}
        
        all_nodes = {**self.junctions, **self.outfalls}
        
        for conduit_name, conduit_data in self.conduits.items():
            from_node = conduit_data['from_node']
            to_node = conduit_data['to_node']
            length = conduit_data['length']
            
            if from_node in all_nodes and to_node in all_nodes:
                from_elevation = all_nodes[from_node]['elevation'] - conduit_data['in_offset']
                to_elevation = all_nodes[to_node]['elevation'] - conduit_data['out_offset']
                
                elevation_diff = from_elevation - to_elevation
                slope = abs(elevation_diff) / length if length > 0 else 0
                slopes[conduit_name] = max(slope, 0.001)  # Ensure positive slope
            else:
                slopes[conduit_name] = 0.001
        
        print(f"Calculated slopes for {len(slopes)} conduits")
        return slopes


class NetworkTopologyAnalyzer:
    """Analyze network topology and calculate CPE2"""
    
    def __init__(self, conduits):
        self.conduits = conduits
        self.network_graph = self._build_network_graph()
        
    def _build_network_graph(self):
        """Build network graph for topology analysis"""
        graph = defaultdict(list)
        reverse_graph = defaultdict(list)
        
        for conduit_name, conduit_data in self.conduits.items():
            from_node = conduit_data['from_node']
            to_node = conduit_data['to_node']
            
            graph[from_node].append((to_node, conduit_name))
            reverse_graph[to_node].append((from_node, conduit_name))
        
        return {'downstream': graph, 'upstream': reverse_graph}
    
    def find_directly_connected_upstream_pipes(self, pipe_name):
        """Find only directly connected upstream pipes"""
        if pipe_name not in self.conduits:
            return []
        
        from_node = self.conduits[pipe_name]['from_node']
        upstream_pipes = []
        
        if from_node in self.network_graph['upstream']:
            for upstream_node, upstream_pipe in self.network_graph['upstream'][from_node]:
                if upstream_pipe != pipe_name:
                    upstream_pipes.append(upstream_pipe)
        
        return upstream_pipes
    
    def find_directly_connected_downstream_pipes(self, pipe_name):
        """Find only directly connected downstream pipes"""
        if pipe_name not in self.conduits:
            return []
        
        to_node = self.conduits[pipe_name]['to_node']
        downstream_pipes = []
        
        if to_node in self.network_graph['downstream']:
            for downstream_node, downstream_pipe in self.network_graph['downstream'][to_node]:
                if downstream_pipe != pipe_name:
                    downstream_pipes.append(downstream_pipe)
        
        return downstream_pipes


def load_diameter_data(excel_file):
    """Load diameter data from Capacity Excel file"""
    print(f"Loading diameter data from {excel_file}")
    
    df = pd.read_excel(excel_file, sheet_name="Capacity")
    df.columns = df.columns.astype(str).str.strip()
    
    # Create diameter dictionary
    diameter_dict = {}
    if ';;Link' in df.columns and 'Geom1(m)' in df.columns:
        for idx, row in df.iterrows():
            if pd.notna(row[';;Link']) and pd.notna(row['Geom1(m)']):
                pipe_id = str(row[';;Link']).strip()
                diameter = float(row['Geom1(m)'])
                diameter_dict[pipe_id] = diameter
    
    print(f"Loaded diameter data for {len(diameter_dict)} pipes")
    return diameter_dict


def calculate_cpe2_for_all_pipes(swmm_parser, diameter_dict):
    """
    Calculate CPE2 (topology-based) for all pipes
    
    CPE2 = upstream_capacity / downstream_capacity
    
    Where capacity = volume * sqrt(slope)
    And volume is calculated from diameter
    """
    print("\n" + "="*80)
    print("CALCULATING CPE2 (TOPOLOGY-BASED)")
    print("="*80)
    print("Formula: CPE2 = Σ(upstream_capacities) / Σ(downstream_capacities)")
    print("="*80)
    
    # Calculate slopes
    slopes = swmm_parser.calculate_slopes()
    
    # Build topology analyzer
    analyzer = NetworkTopologyAnalyzer(swmm_parser.conduits)
    
    # Initialize pipe data storage
    pipe_data = {}
    
    # Calculate capacities for all pipes
    for pipe_name in swmm_parser.conduits.keys():
        pipe_id = str(pipe_name).strip()
        
        # Get diameter
        diameter = diameter_dict.get(pipe_id, 1.0)
        
        # Calculate volume from diameter
        volume = (np.pi * (diameter ** 2) / 4) * 100  # Approximate volume
        
        # Get slope
        slope = slopes.get(pipe_name, 0.001)
        
        # Calculate capacity
        capacity = volume * (slope ** 0.5)
        
        # Ensure capacity is finite
        if not np.isfinite(capacity) or capacity <= 0:
            capacity = volume * 0.1
        
        pipe_data[pipe_id] = {
            'diameter': diameter,
            'slope': slope,
            'capacity': capacity,
        }
    
    # Calculate CPE2 for all pipes
    for pipe_name in swmm_parser.conduits.keys():
        pipe_id = str(pipe_name).strip()
        
        # Find connected pipes
        upstream_pipes = analyzer.find_directly_connected_upstream_pipes(pipe_name)
        downstream_pipes = analyzer.find_directly_connected_downstream_pipes(pipe_name)
        
        # Calculate total upstream capacity
        upstream_capacity = 0
        for up_pipe in upstream_pipes:
            up_pipe_id = str(up_pipe).strip()
            if up_pipe_id in pipe_data:
                upstream_capacity += pipe_data[up_pipe_id]['capacity']
        
        # Calculate total downstream capacity
        downstream_capacity = 0
        for down_pipe in downstream_pipes:
            down_pipe_id = str(down_pipe).strip()
            if down_pipe_id in pipe_data:
                downstream_capacity += pipe_data[down_pipe_id]['capacity']
        
        # Calculate CPE2
        if downstream_capacity > 0:
            cpe2 = upstream_capacity / downstream_capacity
        else:
            cpe2 = 0
        
        # Ensure CPE2 is finite
        if not np.isfinite(cpe2):
            cpe2 = 0
        
        # Store results
        pipe_data[pipe_id]['num_upstream'] = len(upstream_pipes)
        pipe_data[pipe_id]['num_downstream'] = len(downstream_pipes)
        pipe_data[pipe_id]['CPE2'] = cpe2
    
    print(f"\nCPE2 calculated for {len(pipe_data)} pipes")
    
    # Show statistics
    cpe2_values = [data['CPE2'] for data in pipe_data.values() if data['CPE2'] > 0]
    if cpe2_values:
        print(f"CPE2 range: {min(cpe2_values):.4f} to {max(cpe2_values):.4f}")
        print(f"CPE2 mean: {np.mean(cpe2_values):.4f}")
    
    return pipe_data


def save_results(pipe_data, output_file):
    """Save CPE2 results to Excel file"""
    print(f"\nSaving results to {output_file}")
    
    # Create DataFrame
    results = []
    for pipe_id, data in pipe_data.items():
        results.append({
            'Pipe': pipe_id,
            'Diameter': data['diameter'],
            'Slope': data['slope'],
            'Capacity': data['capacity'],
            'Upstream_Pipes': data['num_upstream'],
            'Downstream_Pipes': data['num_downstream'],
            'CPE2': data['CPE2']
        })
    
    df = pd.DataFrame(results)
    
    # Save to Excel
    df.to_excel(output_file, index=False)
    print(f"✅ Results saved successfully!")
    print(f"Total pipes: {len(df)}")


def main():
    """Main execution"""
    print("="*80)
    print("PART 1: SWMM NETWORK ANALYSIS AND CPE2 CALCULATION")
    print("="*80)
    
    # Step 1: Parse SWMM file
    print("\nStep 1: Parsing SWMM network file...")
    swmm_parser = SWMMParser(SWMM_FILE)
    swmm_parser.parse_file()
    
    # Step 2: Load diameter data
    print("\nStep 2: Loading diameter data...")
    diameter_dict = load_diameter_data(CAPACITY_EXCEL)
    
    # Step 3: Calculate CPE2
    print("\nStep 3: Calculating CPE2 (topology-based)...")
    pipe_data = calculate_cpe2_for_all_pipes(swmm_parser, diameter_dict)
    
    # Step 4: Save results
    print("\nStep 4: Saving results...")
    save_results(pipe_data, OUTPUT_FILE)
    
    print("\n" + "="*80)
    print("PART 1 COMPLETED SUCCESSFULLY!")
    print("="*80)
    print(f"Output file: {OUTPUT_FILE}")
    print("\nThis file contains:")
    print("  - Pipe ID")
    print("  - Diameter, Slope, Capacity")
    print("  - Number of upstream/downstream pipes")
    print("  - CPE2 (topology-based capacity performance evaluation)")
    print("\nNext step: Run Part 2 (PVI_Analysis_Main.py) to calculate PVI")


if __name__ == "__main__":
    main()
