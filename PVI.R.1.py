import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

# Set up plotting style (from your original code)
plt.style.use('seaborn-v0_8-whitegrid')

# Define consistent colors and patterns (from your original code)
COLOR_CPE = 'navy'
COLOR_PVI = 'green' 
COLOR_HOURS = 'darkred'

PATTERN_CPE = '///'
PATTERN_PVI = '---'
PATTERN_HOURS = '|||'

def create_visualization(data, title_suffix, filter_condition=None):
    """
    Create visualization for the results using the exact method from your code
    """
    if filter_condition is not None:
        df_filtered = data[filter_condition].copy()
    else:
        df_filtered = data.copy()
    
    # Sort by PVI (descending)
    df_filtered = df_filtered.sort_values('PVI', ascending=False).reset_index(drop=True)
    
    print(f"Creating chart with {len(df_filtered)} pipes: {title_suffix}")
    
    if len(df_filtered) == 0:
        print(f"Warning: No data to plot for {title_suffix}")
        fig, ax = plt.subplots(figsize=(12, 8))
        ax.text(0.5, 0.5, f'No data available for:\n{title_suffix}', 
                horizontalalignment='center', verticalalignment='center',
                transform=ax.transAxes, fontsize=16)
        ax.set_title(f'Weighted CPE Analysis - {title_suffix}', fontsize=18)
        plt.tight_layout()
        return fig
    
    # Create the plot
    fig, ax1 = plt.subplots(figsize=(20, 12), dpi=300)
    
    # Create X-axis positions
    x = np.arange(len(df_filtered))
    
    # Width of each bar group
    width = 0.25
    
    # Plot bars with patterns (using CPE instead of CPE_Weighted)
    bar_cpe = ax1.bar(x - width, df_filtered['CPE'], width, 
                      label='CPE', 
                      color=COLOR_CPE, edgecolor='black', linewidth=1.2,
                      hatch=PATTERN_CPE)
    
    bar_pvi = ax1.bar(x, df_filtered['PVI'], width, 
                      label='PVI', 
                      color=COLOR_PVI, edgecolor='black', linewidth=1.2,
                      hatch=PATTERN_PVI)
    
    # Create second Y-axis for HCL
    ax2 = ax1.twinx()
    
    bar_hours = ax2.bar(x + width, df_filtered['HCL'], width, 
                        label='Hours Capacity Limited', 
                        color=COLOR_HOURS, edgecolor='black', linewidth=1.2,
                        hatch=PATTERN_HOURS)
    
    # Set Y-axis limits
    LEFT_Y_UPPER_LIMIT = 20    # Set your desired upper limit for CPE/PVI axis
    RIGHT_Y_UPPER_LIMIT = 8   # Set your desired upper limit for HCL axis

    # Left Y-axis (CPE and PVI)
    ax1.set_ylim(0, LEFT_Y_UPPER_LIMIT)

    # Right Y-axis (HCL)
    ax2.set_ylim(0, RIGHT_Y_UPPER_LIMIT)
    
    # Set labels and title
    ax1.set_xlabel('Pipe No.', fontsize=22)
    ax1.set_ylabel('CPE or PVI', fontsize=22)
    ax2.set_ylabel('Hours Capacity Limited', fontsize=22)
    
    # Set X-axis ticks and labels
    ax1.set_xticks(x)
    
    if len(df_filtered) > 30:
        ax1.set_xticklabels(df_filtered['Pipe'], rotation=90, fontsize=14)
    else:
        ax1.set_xticklabels(df_filtered['Pipe'], rotation=45, fontsize=18)
    
    # Adding legends
    ax1.legend(loc='upper left', fontsize=18)
    ax2.legend(loc='upper right', fontsize=18)
    
    # Set tick parameters
    ax1.tick_params(axis='both', which='major', labelsize=18)
    ax2.tick_params(axis='both', which='major', labelsize=18)
    
    # Add grid
    ax1.grid(axis='y', linestyle='--', alpha=0.7)
    
    plt.tight_layout()
    return fig

def load_your_data_directly():
    """
    Load your Top_20_PVI data directly into a DataFrame
    """
    data = {
        'Pipe': [338, 186, 235, 36, 34, 210, 99, 335, 154, 200, 208, 174, 303, 314, 319, 125, 311, 98, 132, 309],
        'CPE': [11.4333, 7.586787, 7.415625, 5.458484, 5.443019, 4.400739, 3.843639, 3.841627, 3.006573, 2.884191, 
                2.881871, 2.733168, 2.658869, 2.530972, 2.508719, 2.498437, 2.483846, 2.37296, 2.366854, 2.312981],
        'PVI': [11.4333, 7.586787, 7.415625, 9.454371, 9.427586, 6.023585, 6.657378, 3.841627, 5.207537, 2.884191,
                4.075581, 2.733168, 4.605295, 4.383772, 4.345229, 3.533123, 4.302147, 3.355872, 4.099512, 4.0062],
        'HCL': [5.9, 5.9, 5.9, 5.91, 5.91, 5.89, 5.9, 5.9, 5.82, 5.89, 
                5.87, 5.84, 5.2, 5.33, 5.08, 5.87, 5.24, 5.89, 4.69, 5.42]
    }
    
    df = pd.DataFrame(data)
    print("Your Top_20_PVI Data:")
    print(df)
    print(f"\nData Statistics:")
    print(f"CPE range: {df['CPE'].min():.3f} to {df['CPE'].max():.3f}")
    print(f"PVI range: {df['PVI'].min():.3f} to {df['PVI'].max():.3f}")
    print(f"HCL range: {df['HCL'].min():.3f} to {df['HCL'].max():.3f}")
    
    # Create visualization
    fig = create_visualization(df, "Top 20 PVI - Your Data")
    
    # Save the plot
    output_path = 'Top_20_PVI_Your_Data.png'
    fig.savefig(output_path, dpi=300, bbox_inches='tight')
    print(f"\nVisualization saved to: {output_path}")
    
    # Show the plot
    plt.show()
    
    return df, fig

def load_and_plot_from_excel():
    """
    Load data from Excel file and create visualization
    """
    # File path
    file_path = r"D:\Ph. D\Paper\A topology-based approach for vulnerable pipe identification and modification in the urban drainage network\Publication\ASCE\R1\Figures\Analysis_Results.xlsx"
    
    try:
        # Read the Top_20_PVI sheet
        df = pd.read_excel(file_path, sheet_name='Top_20_PVI')
        
        print(f"Data loaded successfully from Excel!")
        print(f"Shape: {df.shape}")
        print(f"Columns: {df.columns.tolist()}")
        
        # Check if required columns exist
        required_cols = ['Pipe', 'CPE', 'PVI', 'HCL']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            print(f"Missing columns: {missing_cols}")
            print("Available columns:", df.columns.tolist())
            return None, None
        
        # Display first few rows
        print("\nFirst 5 rows of data:")
        print(df[['Pipe', 'CPE', 'PVI', 'HCL']].head())
        
        # Create visualization
        fig = create_visualization(df, "Top 20 PVI from Excel")
        
        # Save the plot
        output_path = os.path.join(os.path.dirname(file_path), 'Top_20_PVI_Visualization.png')
        fig.savefig(output_path, dpi=300, bbox_inches='tight')
        print(f"\nVisualization saved to: {output_path}")
        
        # Show the plot
        plt.show()
        
        return df, fig
        
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        print("Please check the file path and make sure the file exists.")
        return None, None
    except Exception as e:
        print(f"Error loading data: {e}")
        return None, None

def create_custom_visualization(pipe_data, cpe_data, pvi_data, hcl_data):
    """
    Create visualization with custom data arrays
    Use this function if you want to input data manually
    
    Parameters:
    pipe_data: list of pipe names/numbers
    cpe_data: list of CPE values
    pvi_data: list of PVI values  
    hcl_data: list of HCL values
    """
    # Create DataFrame from arrays
    df = pd.DataFrame({
        'Pipe': pipe_data,
        'CPE': cpe_data,
        'PVI': pvi_data,
        'HCL': hcl_data
    })
    
    print(f"Custom data created with {len(df)} pipes")
    
    # Create visualization
    fig = create_visualization(df, "Custom Data")
    
    # Show the plot
    plt.show()
    
    return fig

# Main execution
if __name__ == "__main__":
    print("SWMM Data Visualization Tool")
    print("="*50)
    
    # Option 1: Try to load data from Excel file
    print("Option 1: Attempting to load from Excel file...")
    df_excel, fig_excel = load_and_plot_from_excel()
    
    if df_excel is None:
        print("\nOption 2: Using your provided data directly...")
        df, fig = load_your_data_directly()
    else:
        print("\nExcel data loaded successfully!")
    
    print("\nVisualization complete!")
    print("\nChart Legend:")
    print("- Navy bars with diagonal lines (///) = CPE values (left y-axis)")
    print("- Green bars with horizontal lines (---) = PVI values (left y-axis)") 
    print("- Dark red bars with vertical lines (|||) = HCL values (right y-axis)")
    print("\nThe bars are sorted by PVI values in descending order.")