import os
import glob
import pandas as pd
from datetime import datetime
import dash
from dash import dcc, html, dash_table, Input, Output
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from scipy import stats
from statsmodels.nonparametric.smoothers_lowess import lowess
from scipy.interpolate import griddata, interp1d
import base64
import io

import xml.etree.ElementTree as ET

print("Starting comprehensive POR and TEST data analysis...")

# ===== COATER CONFIGURATION =====
# TEST Data: 8 COATER positions based on wafer order within each XML file (easily configurable)
# Each XML file starts fresh with wafer #1
TEST_COATER_BASE_PATTERN = ['1301', '1401', '1303', '1403', '1302', '1402', '1304', '1404']

# POR Data: COATER values for slots 1, 2, 3 (easily configurable)
POR_COATER_VALUES = ['1', '2', '3']  # Can be changed to any values like ['A', 'B', 'C'] or ['COAT1', 'COAT2', 'COAT3']

# TEST Entity: Hardcoded entity for TEST data (easily configurable)
TEST_ENTITY = 'TZJ501'  # Change this value to update the TEST entity

# DECK Mapping: Group coaters into decks (easily configurable)
COATER_TO_DECK_MAPPING = {
    '1301': '13-L', '1302': '13-L',
    '1303': '13-R', '1304': '13-R', 
    '1401': '14-L', '1402': '14-L',
    '1403': '14-R', '1404': '14-R'
}

# ===== DATA LOADING FUNCTIONS =====

def load_test_data():
    """Load TEST data from XML files"""
    # Directories to search - using TEST directory
    dirs = ['TEST']  # TEST directory for XML files
    patterns = ['*.xml']  # All XML files in the directory

    # Get all files
    files = []
    for d in dirs:
        for pattern in patterns:
            files.extend(glob.glob(os.path.join(d, pattern)))

    print(f"Found {len(files)} XML files to process")

    # Collect data
    records = []
    processed_files = []  # Track successfully processed files

    for file in files:
        try:
            # Get file datetime (last modified)
            file_time = datetime.fromtimestamp(os.path.getmtime(file))
            
            # Extract lot number from filename (between 3rd and 4th dash)
            filename = os.path.basename(file)
            filename_parts = filename.split('-')
            lot_number = filename_parts[3] if len(filename_parts) > 3 else 'Unknown'
            
            # Use hardcoded entity for TEST data (configurable at top of file)
            entity = TEST_ENTITY
            
            # Determine DMT type from file path
            if 'DMT102' in file:
                dmt = 'DMT102'
            elif 'DMT103' in file:
                dmt = 'DMT103'
            elif 'DMT116' in file:
                dmt = 'DMT116'
            else:
                dmt = 'Unknown'
                
            # Parse XML
            tree = ET.parse(file)
            root = tree.getroot()
            
            # Track wafer order within this XML file for COATER assignment
            wafer_order = {}
            wafer_count = 0
            
            # First pass: assign wafer order numbers
            for data_record in root.findall('.//DataRecord'):
                wafer_id = data_record.findtext('WaferID')
                if wafer_id and wafer_id not in wafer_order:
                    wafer_order[wafer_id] = wafer_count
                    wafer_count += 1
            
            # Second pass: process data records with COATER assignment
            # Looking for DataRecord elements with Label and Datum children
            for data_record in root.findall('.//DataRecord'):
                label = data_record.findtext('Label')
                datum = data_record.findtext('Datum')
                wafer_id = data_record.findtext('WaferID')
                x_wafer_loc = data_record.findtext('XWaferLoc')
                y_wafer_loc = data_record.findtext('YWaferLoc')
                slot = data_record.findtext('Slot')
                
                if label in ['Layer 1 Thickness', 'Goodness-of-Fit']:
                    try:
                        datum_val = float(datum)
                        
                        # Round Layer 1 Thickness to 1 decimal place
                        if label == 'Layer 1 Thickness':
                            datum_val = round(datum_val, 1)
                        
                        # Create a unique location identifier for pairing measurements
                        location_id = f"{x_wafer_loc}_{y_wafer_loc}" if x_wafer_loc and y_wafer_loc else None
                        
                        # Calculate RADIUS
                        radius = None
                        if x_wafer_loc and y_wafer_loc:
                            try:
                                x_val = float(x_wafer_loc)
                                y_val = float(y_wafer_loc)
                                radius = np.sqrt(x_val**2 + y_val**2)
                            except (ValueError, TypeError):
                                radius = None
                        
                        # Add COATER column for TEST data based on wafer order within each XML file
                        # Get wafer order (0-indexed) and map to COATER pattern
                        wafer_order_index = wafer_order.get(wafer_id, 0)
                        coater_index = wafer_order_index % len(TEST_COATER_BASE_PATTERN)
                        coater = TEST_COATER_BASE_PATTERN[coater_index]
                        
                        records.append({
                            'datetime': file_time,
                            'Label': label,
                            'Datum': datum_val,
                            'dmt': dmt,
                            'LotNumber': lot_number,
                            'WaferID': wafer_id,
                            'Slot': slot,
                            'COATER': coater,
                            'XWaferLoc': x_wafer_loc,
                            'YWaferLoc': y_wafer_loc,
                            'location_id': location_id,
                            'RADIUS': radius,
                            'Source': 'TEST',
                            'Entity': entity,
                            'FileName': filename  # Add filename for run identification
                        })
                    except (TypeError, ValueError):
                        continue
            
            # Add to processed files list if we got here without errors
            processed_files.append({
                'filename': os.path.basename(file),
                'full_path': file,
                'dmt_type': dmt,
                'lot_number': lot_number,
                'file_datetime': file_time.strftime('%Y-%m-%d %H:%M:%S'),
                'source': 'TEST'
            })
            
        except Exception as e:
            print(f"Error processing file {file}: {e}")
            continue

    return pd.DataFrame(records), processed_files

def load_por_data():
    """Load POR data from CSV file"""
    por_file = 'POR/POR-Raw.csv'
    records = []
    processed_files = []
    
    try:
        # Read the CSV file
        por_df = pd.read_csv(por_file)
        print(f"Loaded POR CSV with {len(por_df)} records and columns: {por_df.columns.tolist()}")
        
        # Process each row
        for _, row in por_df.iterrows():
            try:
                # Extract relevant data
                thickness_val = float(row['CR_VALUE'])  # Already in Angstrom
                x_coord = float(row['X_COORDINATE'])  # Already in mm
                y_coord = float(row['Y_COORDINATE'])  # Already in mm
                
                # Calculate RADIUS from X, Y coordinates
                radius = np.sqrt(x_coord**2 + y_coord**2)
                
                # Parse datetime
                file_time = pd.to_datetime(row['LOT_DATA_COLLECT_DATE'])
                
                # Create location identifier
                location_id = f"{x_coord}_{y_coord}"
                
                records.append({
                    'datetime': file_time,
                    'Label': 'Layer 1 Thickness',  # Standardize the label
                    'Datum': round(thickness_val, 1),
                    'dmt': row.get('ENTITY', 'Unknown'),  # Use ENTITY as equivalent to dmt
                    'LotNumber': row.get('LOT', 'Unknown'),
                    'WaferID': row.get('RAW_WAFER', 'Unknown'),
                    'Slot': row.get('SLOT', 'Unknown'),
                    'COATER': 'Unknown',  # Will be assigned later based on slot ordering
                    'XWaferLoc': str(x_coord),
                    'YWaferLoc': str(y_coord),
                    'location_id': location_id,
                    'RADIUS': radius,
                    'Source': 'POR',
                    'Entity': row.get('ENTITY', 'Unknown')
                })
                
            except (ValueError, TypeError, KeyError) as e:
                print(f"Error processing POR row: {e}")
                continue
        
        # Assign COATER values for POR data based on slot ordering within each run
        # Convert records to DataFrame for easier processing
        temp_df = pd.DataFrame(records)
        if not temp_df.empty:
            # Group by LotNumber and datetime (each run is unique LOT + LOT_DATA_COLLECT_DATE combo)
            temp_df['date_only'] = temp_df['datetime'].dt.date
            
            # Create a list to store updated records
            updated_records = []
            
            for (lot, date), group in temp_df.groupby(['LotNumber', 'date_only']):
                # Sort by slot to ensure proper ordering
                group_sorted = group.sort_values('Slot')
                
                # Assign COATER values: uses configurable POR_COATER_VALUES
                for i, (idx, row) in enumerate(group_sorted.iterrows()):
                    row_dict = row.to_dict()
                    # Remove the temporary date_only column
                    row_dict.pop('date_only', None)
                    # Assign COATER based on position using configurable values
                    if i < len(POR_COATER_VALUES):
                        coater_value = POR_COATER_VALUES[i]
                    else:
                        # If more slots than configured values, use the last configured value
                        coater_value = POR_COATER_VALUES[-1] if POR_COATER_VALUES else 'Unknown'
                    row_dict['COATER'] = coater_value
                    updated_records.append(row_dict)
            
            # Replace the original records list with updated records
            records = updated_records
        
        # Add to processed files list
        processed_files.append({
            'filename': 'POR-Raw.csv',
            'full_path': por_file,
            'dmt_type': 'POR Data',
            'lot_number': 'Multiple',
            'file_datetime': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'source': 'POR'
        })
        
    except Exception as e:
        print(f"Error loading POR data: {e}")
    
    return pd.DataFrame(records), processed_files

# Load both datasets
print("Loading TEST data...")
test_df, test_files = load_test_data()
print(f"TEST data loaded: {len(test_df)} records")

print("Loading POR data...")
por_df, por_files = load_por_data()
print(f"POR data loaded: {len(por_df)} records")

# Combine processed files info
all_processed_files = test_files + por_files

# ===== DATA ANALYSIS FUNCTIONS =====

def normalize_thickness_data(df, method='zscore'):
    """Normalize thickness data for comparison"""
    if df.empty:
        return df
    
    # Only normalize Layer 1 Thickness data
    thickness_data = df[df['Label'] == 'Layer 1 Thickness'].copy()
    other_data = df[df['Label'] != 'Layer 1 Thickness'].copy()
    
    if not thickness_data.empty:
        if method == 'zscore':
            # Z-score normalization: (x - mean) / std
            mean_val = thickness_data['Datum'].mean()
            std_val = thickness_data['Datum'].std()
            thickness_data['Normalized_Datum'] = (thickness_data['Datum'] - mean_val) / std_val
        elif method == 'minmax':
            # Min-Max normalization: (x - min) / (max - min)
            min_val = thickness_data['Datum'].min()
            max_val = thickness_data['Datum'].max()
            thickness_data['Normalized_Datum'] = (thickness_data['Datum'] - min_val) / (max_val - min_val)
        else:  # 'none'
            thickness_data['Normalized_Datum'] = thickness_data['Datum']
    
    # Add normalized column to other data (set to NaN)
    if not other_data.empty:
        other_data['Normalized_Datum'] = np.nan
    
    # Combine back
    result_df = pd.concat([thickness_data, other_data], ignore_index=True)
    return result_df.sort_index()

# Apply normalization
print("Normalizing data...")
if not test_df.empty:
    test_df = normalize_thickness_data(test_df, 'zscore')
if not por_df.empty:
    por_df = normalize_thickness_data(por_df, 'zscore')

# Combine datasets for some analyses
combined_df = pd.concat([test_df, por_df], ignore_index=True) if not test_df.empty and not por_df.empty else test_df if not test_df.empty else por_df

print(f"Combined dataset: {len(combined_df)} records")

# ===== RUN-BASED STATISTICAL ANALYSIS FUNCTIONS =====

def calculate_run_statistics(df):
    """
    Calculate comprehensive run-based statistics for ENTITY/COATER combinations.
    
    For POR: Run = unique LOT_DATA_COLLECT_DATE + COATER
    For TEST: Run = unique FileName + COATER  
    """
    if df.empty:
        return pd.DataFrame()
    
    # Only analyze thickness data
    thickness_df = df[df['Label'] == 'Layer 1 Thickness'].copy()
    
    if thickness_df.empty:
        return pd.DataFrame()
    
    results = []
    
    # Process each entity/coater combination
    for (entity, coater), entity_coater_data in thickness_df.groupby(['Entity', 'COATER']):
        
        # Define runs differently based on data source
        if entity_coater_data['Source'].iloc[0] == 'POR':
            # POR: Run = unique LOT_DATA_COLLECT_DATE + COATER
            run_groups = entity_coater_data.groupby(entity_coater_data['datetime'].dt.date)
        else:
            # TEST: Run = unique FileName + COATER 
            run_groups = entity_coater_data.groupby('FileName')
        
        run_stats = []
        run_means = []
        run_stds = []
        
        # Calculate statistics for each run
        for run_id, run_data in run_groups:
            if len(run_data) > 0:
                run_mean = run_data['Datum'].mean()
                run_std = run_data['Datum'].std() if len(run_data) > 1 else 0.0
                
                run_means.append(run_mean)
                run_stds.append(run_std)
                
                run_stats.append({
                    'run_id': run_id,
                    'mean': run_mean,
                    'std': run_std,
                    'count': len(run_data)
                })
        
        # Calculate the 6 requested statistics if we have enough data
        if len(run_means) >= 1:
            # 1. Mean thickness across runs
            mean_thickness = np.mean(run_means)
            
            # 2. Std dev of mean thickness across runs
            std_mean_thickness = np.std(run_means, ddof=1) if len(run_means) > 1 else 0.0
            
            # 3 & 4. Delta statistics (run-to-run differences)
            if len(run_means) > 1:
                # Sort run_means by time/run order for proper delta calculation
                sorted_means = run_means  # Already in order from groupby
                # Use absolute deltas - positive or negative direction doesn't matter
                deltas = [abs(sorted_means[i+1] - sorted_means[i]) for i in range(len(sorted_means)-1)]
                avg_delta = np.mean(deltas)
                std_delta = np.std(deltas, ddof=1) if len(deltas) > 1 else 0.0
            else:
                avg_delta = 0.0
                std_delta = 0.0
            
            # 5. Average std dev of each run
            avg_run_std = np.mean(run_stds)
            
            # 6. Std dev of the std dev of each run  
            std_run_std = np.std(run_stds, ddof=1) if len(run_stds) > 1 else 0.0
            
            results.append({
                'Entity': entity,
                'COATER': coater,
                'Source': entity_coater_data['Source'].iloc[0],
                'Run_Count': len(run_means),
                'Mean_Thickness': round(mean_thickness, 2),
                'StdDev_Mean_Thickness': round(std_mean_thickness, 3),
                'Avg_Delta_Run_to_Run': round(avg_delta, 3),
                'StdDev_Delta': round(std_delta, 3),
                'Avg_Run_StdDev': round(avg_run_std, 3),
                'StdDev_Run_StdDev': round(std_run_std, 4)
            })
    
    return pd.DataFrame(results)

def create_summary_statistics_table():
    """
    Create a comprehensive summary table comparing all entities and coaters
    """
    # Calculate statistics for combined dataset
    summary_stats = calculate_run_statistics(combined_df)
    
    if summary_stats.empty:
        return html.Div("No data available for summary statistics")
    
    # Sort by Entity and COATER for better organization 
    summary_stats = summary_stats.sort_values(['Entity', 'COATER'])
    
    # Create the dash table
    table = dash_table.DataTable(
        data=summary_stats.to_dict('records'),
        columns=[
            {'name': 'Entity', 'id': 'Entity'},
            {'name': 'COATER', 'id': 'COATER'}, 
            {'name': 'Source', 'id': 'Source'},
            {'name': 'Runs', 'id': 'Run_Count'},
            {'name': 'Mean Thickness (Å)', 'id': 'Mean_Thickness', 'type': 'numeric'},
            {'name': 'StdDev Mean (Å)', 'id': 'StdDev_Mean_Thickness', 'type': 'numeric'},
            {'name': 'Avg Δ Run-to-Run (Å)', 'id': 'Avg_Delta_Run_to_Run', 'type': 'numeric'},
            {'name': 'StdDev Δ (Å)', 'id': 'StdDev_Delta', 'type': 'numeric'},
            {'name': 'Avg Run StdDev (Å)', 'id': 'Avg_Run_StdDev', 'type': 'numeric'},
            {'name': 'StdDev Run StdDev (Å)', 'id': 'StdDev_Run_StdDev', 'type': 'numeric'}
        ],
        style_cell={
            'textAlign': 'center',
            'padding': '10px',
            'fontFamily': 'Arial'
        },
        style_header={
            'backgroundColor': 'rgb(230, 230, 230)',
            'fontWeight': 'bold'
        },
        style_data_conditional=[
            {
                'if': {'filter_query': '{Source} = POR'},
                'backgroundColor': 'rgb(248, 248, 255)',
            },
            {
                'if': {'filter_query': '{Source} = TEST'}, 
                'backgroundColor': 'rgb(255, 248, 248)',
            }
        ],
        sort_action="native",
        page_size=20
    )
    
    return html.Div([
        html.H3("Run-Based Statistical Summary by Entity and COATER"),
        html.P([
            "POR Runs: Unique LOT_DATA_COLLECT_DATE + COATER | ",
            "TEST Runs: Unique FileName + COATER"
        ]),
        table
    ])

def create_entity_comparison_plots():
    """
    Create comparison plots between different entities
    """
    if combined_df.empty:
        return html.Div("No data available for comparison plots")
    
    thickness_data = combined_df[combined_df['Label'] == 'Layer 1 Thickness']
    
    if thickness_data.empty:
        return html.Div("No thickness data available for comparison")
    
    # Box plot comparing entities by coater
    fig1 = px.box(
        thickness_data,
        x='COATER',
        y='Datum',
        color='Entity',
        facet_col='Source',
        title='Thickness Distribution by Entity and COATER',
        labels={'Datum': 'Layer 1 Thickness (Å)', 'COATER': 'COATER'}
    )
    
    # Mean thickness comparison
    entity_stats = thickness_data.groupby(['Entity', 'COATER', 'Source'])['Datum'].agg(['mean', 'std']).reset_index()
    
    fig2 = px.bar(
        entity_stats,
        x='COATER',
        y='mean',
        color='Entity',
        facet_col='Source',
        error_y='std',
        title='Mean Thickness Comparison by Entity and COATER',
        labels={'mean': 'Mean Thickness (Å)', 'COATER': 'COATER'}
    )
    
    return html.Div([
        html.H3("Entity Comparison Plots"),
        dcc.Graph(figure=fig1),
        dcc.Graph(figure=fig2)
    ])

# ===== VISUALIZATION FUNCTIONS =====

def apply_offset_to_test_data(test_data, offset):
    """Apply matching offset to TEST thickness data"""
    if test_data.empty or offset == 0:
        return test_data
    
    modified_data = test_data.copy()
    # Apply offset only to Layer 1 Thickness measurements
    mask = modified_data['Label'] == 'Layer 1 Thickness'
    modified_data.loc[mask, 'Datum'] = modified_data.loc[mask, 'Datum'] + offset
    
    # Also update normalized data if it exists
    if 'Normalized_Datum' in modified_data.columns:
        # Recalculate normalized values with the offset applied
        thickness_data = modified_data[mask]
        if not thickness_data.empty:
            mean_val = thickness_data['Datum'].mean()
            std_val = thickness_data['Datum'].std()
            modified_data.loc[mask, 'Normalized_Datum'] = (thickness_data['Datum'] - mean_val) / std_val
    
    return modified_data

def make_source_comparison_boxplot(offset=0):
    """Create boxplot comparing TEST vs POR thickness data with applied offset"""
    # Apply offset to TEST data
    modified_test_df = apply_offset_to_test_data(test_df, offset)
    modified_combined_df = pd.concat([modified_test_df, por_df], ignore_index=True) if not modified_test_df.empty and not por_df.empty else modified_test_df if not modified_test_df.empty else por_df
    
    if modified_combined_df.empty:
        return html.Div("No data available for comparison")
    
    # Filter for thickness data only
    thickness_data = combined_df[combined_df['Label'] == 'Layer 1 Thickness']
    
    if thickness_data.empty:
        return html.Div("No thickness data available")
    
    # Filter for thickness data only
    thickness_data = modified_combined_df[modified_combined_df['Label'] == 'Layer 1 Thickness']
    
    if thickness_data.empty:
        return html.Div("No thickness data available")
    
    # Add offset info to title
    title_text = f'Thickness Data Comparison: TEST vs POR (TEST Offset: {offset:+.1f} Å)' if offset != 0 else 'Thickness Data Comparison: TEST vs POR'
    
    fig = px.box(
        thickness_data,
        x='Source',
        y='Datum',
        color='Entity',
        points='all',
        title=title_text,
        labels={'Datum': 'Layer 1 Thickness (Angstrom)', 'Source': 'Data Source'}
    )
    
    fig.update_layout(
        xaxis_title='Data Source',
        yaxis_title='Layer 1 Thickness (Angstrom)',
        height=600
    )
    
    return dcc.Graph(figure=fig)

def make_normalized_comparison_boxplot(offset=0):
    """Create boxplot comparing normalized TEST vs POR thickness data with applied offset"""
    # Apply offset to TEST data
    modified_test_df = apply_offset_to_test_data(test_df, offset)
    modified_combined_df = pd.concat([modified_test_df, por_df], ignore_index=True) if not modified_test_df.empty and not por_df.empty else modified_test_df if not modified_test_df.empty else por_df
    
    if modified_combined_df.empty:
        return html.Div("No data available for normalized comparison")
    
    # Filter for thickness data with normalized values
    thickness_data = modified_combined_df[(modified_combined_df['Label'] == 'Layer 1 Thickness') & 
                                (modified_combined_df['Normalized_Datum'].notna())]
    
    if thickness_data.empty:
        return html.Div("No normalized thickness data available")
    
    # Add offset info to title
    title_text = f'Normalized Thickness Data Comparison: TEST vs POR (TEST Offset: {offset:+.1f} Å)' if offset != 0 else 'Normalized Thickness Data Comparison: TEST vs POR'
    
    fig = px.box(
        thickness_data,
        x='Source',
        y='Normalized_Datum',
        color='Entity',
        points='all',
        title=title_text,
        labels={'Normalized_Datum': 'Normalized Layer 1 Thickness (Z-Score)', 'Source': 'Data Source'}
    )
    
    fig.update_layout(
        xaxis_title='Data Source',
        yaxis_title='Normalized Layer 1 Thickness (Z-Score)',
        height=600
    )
    
    return dcc.Graph(figure=fig)

def make_statistics_comparison_table(offset=0):
    """Create comprehensive run-based statistics comparison table with applied offset"""
    # Apply offset to TEST data
    modified_test_df = apply_offset_to_test_data(test_df, offset)
    modified_combined_df = pd.concat([modified_test_df, por_df], ignore_index=True) if not modified_test_df.empty and not por_df.empty else modified_test_df if not modified_test_df.empty else por_df
    
    if modified_combined_df.empty:
        return html.Div("No data available for statistics")
    
    # Calculate run-based statistics using the modified data
    summary_stats = calculate_run_statistics(modified_combined_df)
    
    if summary_stats.empty:
        return html.Div("No run-based statistics available")
    
    # Sort by Entity and COATER for better organization 
    summary_stats = summary_stats.sort_values(['Entity', 'COATER'])
    
    # Create the comprehensive dash table
    table = dash_table.DataTable(
        data=summary_stats.to_dict('records'),
        columns=[
            {'name': 'Entity', 'id': 'Entity', 'type': 'text'},
            {'name': 'COATER', 'id': 'COATER', 'type': 'text'}, 
            {'name': 'Source', 'id': 'Source', 'type': 'text'},
            {'name': 'Runs', 'id': 'Run_Count', 'type': 'numeric'},
            {'name': 'Mean Thickness (Å)', 'id': 'Mean_Thickness', 'type': 'numeric'},
            {'name': 'StdDev Mean (Å)', 'id': 'StdDev_Mean_Thickness', 'type': 'numeric'},
            {'name': 'Avg Δ Run-to-Run (Å)', 'id': 'Avg_Delta_Run_to_Run', 'type': 'numeric'},
            {'name': 'StdDev Δ (Å)', 'id': 'StdDev_Delta', 'type': 'numeric'},
            {'name': 'Avg Run StdDev (Å)', 'id': 'Avg_Run_StdDev', 'type': 'numeric'},
            {'name': 'StdDev Run StdDev (Å)', 'id': 'StdDev_Run_StdDev', 'type': 'numeric'}
        ],
        style_cell={
            'textAlign': 'center',
            'padding': '8px',
            'fontFamily': 'Arial',
            'fontSize': '12px'
        },
        style_header={
            'backgroundColor': 'rgb(230, 230, 230)',
            'fontWeight': 'bold',
            'fontSize': '13px'
        },
        style_data_conditional=[
            {
                'if': {'filter_query': '{Source} = POR'},
                'backgroundColor': 'rgb(248, 248, 255)',
            },
            {
                'if': {'filter_query': '{Source} = TEST'}, 
                'backgroundColor': 'rgb(255, 248, 248)',
            }
        ],
        sort_action="native",
        page_size=20,
        style_table={'overflowX': 'auto'}
    )
    
    # Add offset info to title
    title_text = f"Run-Based Statistical Summary by Entity and COATER (TEST Offset: {offset:+.1f} Å)" if offset != 0 else "Run-Based Statistical Summary by Entity and COATER"
    
    description = html.Div([
        html.P([
            html.Strong("Run Definitions: "),
            "POR = Unique LOT_DATA_COLLECT_DATE + COATER | ",
            "TEST = Unique FileName + COATER"
        ]),
        html.P([
            html.Strong("Statistics: "),
            "6 metrics per Entity/COATER: Mean, StdDev of means, Avg delta run-to-run, ",
            "StdDev of deltas, Avg within-run StdDev, StdDev of within-run StdDevs"
        ])
    ], style={'fontSize': '12px', 'marginBottom': '15px', 'fontStyle': 'italic'})
    
    return html.Div([
        html.H3(title_text),
        description,
        table
    ])

def make_radius_spline_plots(offset=0):
    """Create spline plots from center to edge for both TEST and POR with applied offset"""
    # Apply offset to TEST data
    modified_test_df = apply_offset_to_test_data(test_df, offset)
    modified_combined_df = pd.concat([modified_test_df, por_df], ignore_index=True) if not modified_test_df.empty and not por_df.empty else modified_test_df if not modified_test_df.empty else por_df
    
    if modified_combined_df.empty:
        return html.Div("No data available for spline plots")
    
    # Filter for thickness data with valid radius
    thickness_data = modified_combined_df[(modified_combined_df['Label'] == 'Layer 1 Thickness') & 
                                (modified_combined_df['RADIUS'].notna())].copy()
    
    if thickness_data.empty:
        return html.Div("No thickness data with radius available")
    
    fig = go.Figure()
    
    # Define colors for different sources and entities
    color_map = {
        'TEST_TZJ501': '#1f77b4',
        'POR_TTB134': '#ff7f0e',
        'POR_Others': '#2ca02c'
    }
    
    # Process each source separately
    for source in ['TEST', 'POR']:
        source_data = thickness_data[thickness_data['Source'] == source]
        
        if source_data.empty:
            continue
        
        # Group by wafer/entity for individual splines
        if source == 'TEST':
            # Group by WaferID for TEST data
            for wafer_id in source_data['WaferID'].unique():
                wafer_data = source_data[source_data['WaferID'] == wafer_id].copy()
                
                if len(wafer_data) < 3:  # Need at least 3 points for spline
                    continue
                
                # Sort by radius for proper spline
                wafer_data = wafer_data.sort_values('RADIUS')
                
                # Create spline interpolation
                try:
                    # Filter data to 0-150mm radius range
                    wafer_data = wafer_data[wafer_data['RADIUS'] <= 150]
                    
                    if len(wafer_data) >= 3:
                        # Create interpolation function
                        f = interp1d(wafer_data['RADIUS'], wafer_data['Datum'], 
                                   kind='cubic', bounds_error=False, fill_value='extrapolate')
                        
                        # Generate smooth curve from 0 to 150mm
                        radius_smooth = np.linspace(0, 150, 100)
                        thickness_smooth = f(radius_smooth)
                        
                        # Add trace for this wafer
                        fig.add_trace(go.Scatter(
                            x=radius_smooth,
                            y=thickness_smooth,
                            mode='lines',
                            name=f'TEST - {wafer_id}',
                            line=dict(color=color_map.get('TEST_TZJ501', '#1f77b4'), width=2),
                            opacity=0.7,
                            hovertemplate=f'<b>TEST - {wafer_id}</b><br>' +
                                          'Radius: %{x:.1f} mm<br>' +
                                          'Thickness: %{y:.1f} Å<br>' +
                                          '<extra></extra>'
                        ))
                        
                        # Add actual data points
                        fig.add_trace(go.Scatter(
                            x=wafer_data['RADIUS'],
                            y=wafer_data['Datum'],
                            mode='markers',
                            name=f'TEST Points - {wafer_id}',
                            marker=dict(color=color_map.get('TEST_TZJ501', '#1f77b4'), size=6),
                            showlegend=False,
                            hovertemplate=f'<b>TEST Data - {wafer_id}</b><br>' +
                                          'Radius: %{x:.1f} mm<br>' +
                                          'Thickness: %{y:.1f} Å<br>' +
                                          '<extra></extra>'
                        ))
                        
                except Exception as e:
                    print(f"Error creating spline for TEST wafer {wafer_id}: {e}")
                    continue
        
        else:  # POR data
            # Group by Entity for POR data
            for entity in source_data['Entity'].unique():
                entity_data = source_data[source_data['Entity'] == entity].copy()
                
                if len(entity_data) < 3:  # Need at least 3 points for spline
                    continue
                
                # Sort by radius for proper spline
                entity_data = entity_data.sort_values('RADIUS')
                
                # Create spline interpolation
                try:
                    # Filter data to 0-150mm radius range
                    entity_data = entity_data[entity_data['RADIUS'] <= 150]
                    
                    if len(entity_data) >= 3:
                        # Create interpolation function
                        f = interp1d(entity_data['RADIUS'], entity_data['Datum'], 
                                   kind='cubic', bounds_error=False, fill_value='extrapolate')
                        
                        # Generate smooth curve from 0 to 150mm
                        radius_smooth = np.linspace(0, 150, 100)
                        thickness_smooth = f(radius_smooth)
                        
                        # Determine color based on entity
                        color_key = 'POR_TTB134' if entity == 'TTB134' else 'POR_Others'
                        color = color_map.get(color_key, '#2ca02c')
                        
                        # Add trace for this entity
                        fig.add_trace(go.Scatter(
                            x=radius_smooth,
                            y=thickness_smooth,
                            mode='lines',
                            name=f'POR - {entity}',
                            line=dict(color=color, width=2),
                            opacity=0.7,
                            hovertemplate=f'<b>POR - {entity}</b><br>' +
                                          'Radius: %{x:.1f} mm<br>' +
                                          'Thickness: %{y:.1f} Å<br>' +
                                          '<extra></extra>'
                        ))
                        
                        # Add actual data points
                        fig.add_trace(go.Scatter(
                            x=entity_data['RADIUS'],
                            y=entity_data['Datum'],
                            mode='markers',
                            name=f'POR Points - {entity}',
                            marker=dict(color=color, size=6),
                            showlegend=False,
                            hovertemplate=f'<b>POR Data - {entity}</b><br>' +
                                          'Radius: %{x:.1f} mm<br>' +
                                          'Thickness: %{y:.1f} Å<br>' +
                                          '<extra></extra>'
                        ))
                        
                except Exception as e:
                    print(f"Error creating spline for POR entity {entity}: {e}")
                    continue
    
    # Update layout
    title_text = f'Radial Thickness Profiles: TEST vs POR (0-150mm, TEST Offset: {offset:+.1f} Å)' if offset != 0 else 'Radial Thickness Profiles: TEST vs POR (0-150mm)'
    fig.update_layout(
        title=title_text,
        xaxis_title='Radius from Center (mm)',
        yaxis_title='Layer 1 Thickness (Angstrom)',
        xaxis=dict(range=[0, 150], dtick=25),
        height=700,
        margin=dict(l=50, r=50, t=50, b=50),
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        )
    )
    
    return dcc.Graph(figure=fig)

def make_coordinate_comparison_scatter(offset=0):
    """Create scatter plot comparing thickness at matching X,Y coordinates with applied offset"""
    # Apply offset to TEST data
    modified_test_df = apply_offset_to_test_data(test_df, offset)
    
    if por_df.empty or modified_test_df.empty:
        return html.Div("Need both TEST and POR data for coordinate comparison")
    
    # Filter for thickness data
    test_thickness = modified_test_df[modified_test_df['Label'] == 'Layer 1 Thickness'].copy()
    por_thickness = por_df[por_df['Label'] == 'Layer 1 Thickness'].copy()
    
    if test_thickness.empty or por_thickness.empty:
        return html.Div("No thickness data available for comparison")
    
    # Round coordinates to find matches (within 1mm tolerance)
    test_thickness['X_Round'] = test_thickness['XWaferLoc'].astype(float).round(0)
    test_thickness['Y_Round'] = test_thickness['YWaferLoc'].astype(float).round(0)
    por_thickness['X_Round'] = por_thickness['XWaferLoc'].astype(float).round(0)
    por_thickness['Y_Round'] = por_thickness['YWaferLoc'].astype(float).round(0)
    
    # Merge on rounded coordinates to find matches
    merged_data = pd.merge(
        test_thickness[['X_Round', 'Y_Round', 'Datum', 'Normalized_Datum', 'WaferID']].rename(columns={
            'Datum': 'TEST_Thickness',
            'Normalized_Datum': 'TEST_Normalized',
            'WaferID': 'TEST_WaferID'
        }),
        por_thickness[['X_Round', 'Y_Round', 'Datum', 'Normalized_Datum', 'Entity']].rename(columns={
            'Datum': 'POR_Thickness',
            'Normalized_Datum': 'POR_Normalized',
            'Entity': 'POR_Entity'
        }),
        on=['X_Round', 'Y_Round'],
        how='inner'
    )
    
    if merged_data.empty:
        return html.Div("No matching coordinate locations found between TEST and POR data")
    
    # Create scatter plot
    fig = go.Figure()
    
    # Raw thickness comparison
    fig.add_trace(go.Scatter(
        x=merged_data['TEST_Thickness'],
        y=merged_data['POR_Thickness'],
        mode='markers',
        name='Raw Thickness',
        marker=dict(size=8, color='blue', opacity=0.6),
        text=merged_data.apply(lambda row: f"X: {row['X_Round']}, Y: {row['Y_Round']}<br>"
                                           f"TEST: {row['TEST_WaferID']}<br>"
                                           f"POR: {row['POR_Entity']}", axis=1),
        hovertemplate='<b>Coordinate Match</b><br>' +
                      '%{text}<br>' +
                      'TEST Thickness: %{x:.1f} Å<br>' +
                      'POR Thickness: %{y:.1f} Å<br>' +
                      '<extra></extra>'
    ))
    
    # Add diagonal reference line (perfect correlation)
    min_val = min(merged_data['TEST_Thickness'].min(), merged_data['POR_Thickness'].min())
    max_val = max(merged_data['TEST_Thickness'].max(), merged_data['POR_Thickness'].max())
    
    fig.add_trace(go.Scatter(
        x=[min_val, max_val],
        y=[min_val, max_val],
        mode='lines',
        name='Perfect Correlation',
        line=dict(color='red', dash='dash', width=2),
        hoverinfo='skip'
    ))
    
    # Calculate correlation
    correlation = merged_data['TEST_Thickness'].corr(merged_data['POR_Thickness'])
    
    # Add offset info to title
    title_text = f'Thickness Comparison at Matching Coordinates (r = {correlation:.3f}, TEST Offset: {offset:+.1f} Å)' if offset != 0 else f'Thickness Comparison at Matching Coordinates (r = {correlation:.3f})'
    
    fig.update_layout(
        title=title_text,
        xaxis_title='TEST Thickness (Angstrom)',
        yaxis_title='POR Thickness (Angstrom)',
        height=600,
        margin=dict(l=50, r=50, t=80, b=50),
        showlegend=True
    )
    
    return dcc.Graph(figure=fig)

def make_test_time_series_plot(offset=0):
    """Create time series plot of TEST thickness data over time with applied offset"""
    # Apply offset to TEST data
    modified_test_df = apply_offset_to_test_data(test_df, offset)
    
    if modified_test_df.empty:
        return html.Div("No TEST data available for time series analysis")
    
    # Filter for thickness data only
    thickness_data = modified_test_df[modified_test_df['Label'] == 'Layer 1 Thickness'].copy()
    
    if thickness_data.empty:
        return html.Div("No TEST thickness data available")
    
    # Sort by datetime for proper time series
    thickness_data = thickness_data.sort_values('datetime')
    
    # Create the time series plot
    fig = go.Figure()
    
    # Group by COATER and calculate mean thickness for each file time point
    # Each file time corresponds to a unique measurement session
    
    # Group by COATER to show different coater traces
    for coater in sorted(thickness_data['COATER'].unique()):
        coater_data = thickness_data[thickness_data['COATER'] == coater]
        
        # Calculate mean thickness for each datetime (file) for this coater
        file_means = coater_data.groupby('datetime').agg({
            'Datum': 'mean',
            'WaferID': 'count',    # Count of wafers for this coater at this file time
            'FileName': 'first'    # Get filename for reference
        }).reset_index()
        
        # Sort by datetime
        file_means = file_means.sort_values('datetime')
        
        # Create categorical x-axis labels (file times)
        x_labels = [dt.strftime('%Y-%m-%d %H:%M:%S') for dt in file_means['datetime']]
        
        # Add trace for each COATER (showing mean values at each file time)
        fig.add_trace(go.Scatter(
            x=x_labels,  # Use formatted datetime strings as categories
            y=file_means['Datum'],
            mode='markers+lines',
            name=f'COATER {coater}',
            marker=dict(size=10),
            line=dict(width=3),
            hovertemplate='<b>COATER %{fullData.name}</b><br>' +
                          'File Time: %{x}<br>' +
                          'Mean Thickness: %{y:.1f} Å<br>' +
                          'Wafer Count: %{customdata[0]}<br>' +
                          'File: %{customdata[1]}<br>' +
                          '<extra></extra>',
            customdata=file_means[['WaferID', 'FileName']].values
        ))
    
    # Add offset info to title
    title_text = f'TEST Thickness Over Time (Offset: {offset:+.1f} Å)' if offset != 0 else 'TEST Thickness Over Time'
    
    # Update layout
    fig.update_layout(
        title=title_text,
        xaxis_title='File Date/Time',
        yaxis_title='Layer 1 Thickness (Angstrom)',
        xaxis=dict(
            type='category',  # Treat x-axis as categorical
            tickangle=45,     # Rotate labels for better readability
            showgrid=True
        ),
        yaxis=dict(showgrid=True),
        height=600,
        margin=dict(l=50, r=50, t=80, b=100),  # Extra bottom margin for rotated labels
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        )
    )
    
    return dcc.Graph(figure=fig)

def make_test_deck_time_series_plot(offset=0):
    """Create time series plot of TEST thickness data grouped by DECK with applied offset"""
    # Apply offset to TEST data
    modified_test_df = apply_offset_to_test_data(test_df, offset)
    
    if modified_test_df.empty:
        return html.Div("No TEST data available for deck time series analysis")
    
    # Filter for thickness data only
    thickness_data = modified_test_df[modified_test_df['Label'] == 'Layer 1 Thickness'].copy()
    
    if thickness_data.empty:
        return html.Div("No TEST thickness data available")
    
    # Add DECK column based on COATER mapping
    thickness_data['DECK'] = thickness_data['COATER'].map(COATER_TO_DECK_MAPPING)
    
    # Remove any data that doesn't map to a deck
    thickness_data = thickness_data.dropna(subset=['DECK'])
    
    if thickness_data.empty:
        return html.Div("No TEST thickness data with valid DECK mapping")
    
    # Sort by datetime for proper time series
    thickness_data = thickness_data.sort_values('datetime')
    
    # Create the time series plot
    fig = go.Figure()
    
    # Group by DECK and calculate mean thickness for each file time point
    for deck in sorted(thickness_data['DECK'].unique()):
        deck_data = thickness_data[thickness_data['DECK'] == deck]
        
        # Calculate mean thickness for each datetime (file) for this deck
        file_means = deck_data.groupby('datetime').agg({
            'Datum': 'mean',
            'WaferID': 'count',    # Count of wafers for this deck at this file time
            'FileName': 'first',   # Get filename for reference
            'COATER': lambda x: ', '.join(sorted(x.unique()))  # Show which coaters contributed
        }).reset_index()
        
        # Sort by datetime
        file_means = file_means.sort_values('datetime')
        
        # Create categorical x-axis labels (file times)
        x_labels = [dt.strftime('%Y-%m-%d %H:%M:%S') for dt in file_means['datetime']]
        
        # Add trace for each DECK (showing mean values at each file time)
        fig.add_trace(go.Scatter(
            x=x_labels,  # Use formatted datetime strings as categories
            y=file_means['Datum'],
            mode='markers+lines',
            name=f'DECK {deck}',
            marker=dict(size=12),  # Slightly larger for deck view
            line=dict(width=4),    # Thicker lines for deck view
            hovertemplate='<b>DECK %{fullData.name}</b><br>' +
                          'File Time: %{x}<br>' +
                          'Mean Thickness: %{y:.1f} Å<br>' +
                          'Wafer Count: %{customdata[0]}<br>' +
                          'COATERs: %{customdata[1]}<br>' +
                          'File: %{customdata[2]}<br>' +
                          '<extra></extra>',
            customdata=file_means[['WaferID', 'COATER', 'FileName']].values
        ))
    
    # Add offset info to title
    title_text = f'TEST Thickness by DECK Over Time (Offset: {offset:+.1f} Å)' if offset != 0 else 'TEST Thickness by DECK Over Time'
    
    # Update layout
    fig.update_layout(
        title=title_text,
        xaxis_title='File Date/Time',
        yaxis_title='Layer 1 Thickness (Angstrom)',
        xaxis=dict(
            type='category',  # Treat x-axis as categorical
            tickangle=45,     # Rotate labels for better readability
            showgrid=True
        ),
        yaxis=dict(showgrid=True),
        height=600,
        margin=dict(l=50, r=50, t=80, b=100),  # Extra bottom margin for rotated labels
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        )
    )
    
    return dcc.Graph(figure=fig)

def make_test_radius_spline_plot(selected_files=None, offset=0):
    """Create spline plot of TEST thickness vs radius by DECK for selected files"""
    from scipy.interpolate import UnivariateSpline
    
    # Apply offset to TEST data
    modified_test_df = apply_offset_to_test_data(test_df, offset)
    
    if modified_test_df.empty:
        return html.Div("No TEST data available for spline analysis")
    
    # Filter for thickness data only
    thickness_data = modified_test_df[modified_test_df['Label'] == 'Layer 1 Thickness'].copy()
    
    if thickness_data.empty:
        return html.Div("No TEST thickness data available")
    
    # Add DECK mapping
    thickness_data['DECK'] = thickness_data['COATER'].map(COATER_TO_DECK_MAPPING)
    
    # Filter by selected files if specified
    if selected_files:
        thickness_data = thickness_data[thickness_data['FileName'].isin(selected_files)]
        
    if thickness_data.empty:
        return html.Div("No data available for selected files")
    
    # Filter for valid radius data (0-150mm)
    thickness_data = thickness_data[
        (thickness_data['RADIUS'].notna()) & 
        (thickness_data['RADIUS'] >= 0) & 
        (thickness_data['RADIUS'] <= 150)
    ].copy()
    
    if thickness_data.empty:
        return html.Div("No thickness data with valid radius (0-150mm) available")
    
    # Create the spline plot
    fig = go.Figure()
    
    # Define colors for each DECK
    deck_colors = {
        '13-L': '#1f77b4',  # Blue
        '13-R': '#ff7f0e',  # Orange
        '14-L': '#2ca02c',  # Green
        '14-R': '#d62728'   # Red
    }
    
    # Create splines for each DECK
    for deck in sorted(thickness_data['DECK'].unique()):
        deck_data = thickness_data[thickness_data['DECK'] == deck]
        
        if len(deck_data) < 4:  # Need at least 4 points for spline
            continue
            
        # Sort by radius for proper spline fitting
        deck_data = deck_data.sort_values('RADIUS')
        
        # Create spline
        try:
            # Use moderate smoothing (s parameter)
            spline = UnivariateSpline(deck_data['RADIUS'], deck_data['Datum'], s=len(deck_data)*2)
            
            # Generate smooth curve points
            radius_smooth = np.linspace(deck_data['RADIUS'].min(), deck_data['RADIUS'].max(), 100)
            thickness_smooth = spline(radius_smooth)
            
            # Add raw data points
            fig.add_trace(go.Scatter(
                x=deck_data['RADIUS'],
                y=deck_data['Datum'],
                mode='markers',
                name=f'{deck} Data',
                marker=dict(
                    size=8,
                    color=deck_colors.get(deck, '#666666'),
                    opacity=0.6
                ),
                hovertemplate=f'<b>DECK {deck} Data</b><br>' +
                              'Radius: %{x:.1f} mm<br>' +
                              'Thickness: %{y:.1f} Å<br>' +
                              '<extra></extra>'
            ))
            
            # Add spline curve
            fig.add_trace(go.Scatter(
                x=radius_smooth,
                y=thickness_smooth,
                mode='lines',
                name=f'{deck} Spline',
                line=dict(
                    width=3,
                    color=deck_colors.get(deck, '#666666')
                ),
                hovertemplate=f'<b>DECK {deck} Spline</b><br>' +
                              'Radius: %{x:.1f} mm<br>' +
                              'Thickness: %{y:.1f} Å<br>' +
                              '<extra></extra>'
            ))
            
        except Exception as e:
            # If spline fails, just show raw data
            fig.add_trace(go.Scatter(
                x=deck_data['RADIUS'],
                y=deck_data['Datum'],
                mode='markers+lines',
                name=f'{deck} (Linear)',
                marker=dict(size=8, color=deck_colors.get(deck, '#666666')),
                line=dict(width=2, color=deck_colors.get(deck, '#666666')),
                hovertemplate=f'<b>DECK {deck}</b><br>' +
                              'Radius: %{x:.1f} mm<br>' +
                              'Thickness: %{y:.1f} Å<br>' +
                              '<extra></extra>'
            ))
    
    # Create title with file info
    file_info = f" ({len(selected_files)} files selected)" if selected_files else " (All files)"
    title_text = f'TEST Thickness vs Radius by DECK{file_info}'
    if offset != 0:
        title_text += f' (Offset: {offset:+.1f} Å)'
    
    # Update layout
    fig.update_layout(
        title=title_text,
        xaxis_title='Radius from Center (mm)',
        yaxis_title='Layer 1 Thickness (Angstrom)',
        xaxis=dict(
            range=[0, 150],
            showgrid=True
        ),
        yaxis=dict(showgrid=True),
        height=600,
        margin=dict(l=50, r=50, t=80, b=50),
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        )
    )
    
    return dcc.Graph(figure=fig)

def make_por_radius_spline_plot(selected_entities=None):
    """Create spline plot of POR thickness vs radius by COATER for selected entities"""
    from scipy.interpolate import UnivariateSpline
    
    if por_df.empty:
        return html.Div("No POR data available for spline analysis")
    
    # Filter for thickness data only
    thickness_data = por_df[por_df['Label'] == 'Layer 1 Thickness'].copy()
    
    if thickness_data.empty:
        return html.Div("No POR thickness data available")
    
    # Filter by selected entities if specified
    if selected_entities:
        thickness_data = thickness_data[thickness_data['Entity'].isin(selected_entities)]
        
    if thickness_data.empty:
        return html.Div("No data available for selected entities")
    
    # Filter for valid radius data (0-150mm)
    thickness_data = thickness_data[
        (thickness_data['RADIUS'].notna()) & 
        (thickness_data['RADIUS'] >= 0) & 
        (thickness_data['RADIUS'] <= 150)
    ].copy()
    
    if thickness_data.empty:
        return html.Div("No thickness data with valid radius (0-150mm) available")
    
    # Create the spline plot
    fig = go.Figure()
    
    # Define colors for each COATER
    coater_colors = {
        '1': '#1f77b4',  # Blue
        '2': '#ff7f0e',  # Orange
        '3': '#2ca02c',  # Green
    }
    
    # Create splines for each COATER
    for coater in sorted(thickness_data['COATER'].unique()):
        coater_data = thickness_data[thickness_data['COATER'] == coater]
        
        if len(coater_data) < 4:  # Need at least 4 points for spline
            continue
            
        # Sort by radius for proper spline fitting
        coater_data = coater_data.sort_values('RADIUS')
        
        # Create spline
        try:
            # Use moderate smoothing (s parameter)
            spline = UnivariateSpline(coater_data['RADIUS'], coater_data['Datum'], s=len(coater_data)*2)
            
            # Generate smooth curve points
            radius_smooth = np.linspace(coater_data['RADIUS'].min(), coater_data['RADIUS'].max(), 100)
            thickness_smooth = spline(radius_smooth)
            
            # Add raw data points
            fig.add_trace(go.Scatter(
                x=coater_data['RADIUS'],
                y=coater_data['Datum'],
                mode='markers',
                name=f'COATER {coater} Data',
                marker=dict(
                    size=8,
                    color=coater_colors.get(str(coater), '#666666'),
                    opacity=0.6
                ),
                hovertemplate=f'<b>COATER {coater} Data</b><br>' +
                              'Radius: %{x:.1f} mm<br>' +
                              'Thickness: %{y:.1f} Å<br>' +
                              '<extra></extra>'
            ))
            
            # Add spline curve
            fig.add_trace(go.Scatter(
                x=radius_smooth,
                y=thickness_smooth,
                mode='lines',
                name=f'COATER {coater} Spline',
                line=dict(
                    width=3,
                    color=coater_colors.get(str(coater), '#666666')
                ),
                hovertemplate=f'<b>COATER {coater} Spline</b><br>' +
                              'Radius: %{x:.1f} mm<br>' +
                              'Thickness: %{y:.1f} Å<br>' +
                              '<extra></extra>'
            ))
            
        except Exception as e:
            # If spline fails, just show raw data
            fig.add_trace(go.Scatter(
                x=coater_data['RADIUS'],
                y=coater_data['Datum'],
                mode='markers+lines',
                name=f'COATER {coater} (Linear)',
                marker=dict(size=8, color=coater_colors.get(str(coater), '#666666')),
                line=dict(width=2, color=coater_colors.get(str(coater), '#666666')),
                hovertemplate=f'<b>COATER {coater}</b><br>' +
                              'Radius: %{x:.1f} mm<br>' +
                              'Thickness: %{y:.1f} Å<br>' +
                              '<extra></extra>'
            ))
    
    # Create title with entity info
    entity_info = f" ({len(selected_entities)} entities selected)" if selected_entities else " (All entities)"
    title_text = f'POR Thickness vs Radius by COATER{entity_info}'
    
    # Update layout
    fig.update_layout(
        title=title_text,
        xaxis_title='Radius from Center (mm)',
        yaxis_title='Layer 1 Thickness (Angstrom)',
        xaxis=dict(
            range=[0, 150],
            showgrid=True
        ),
        yaxis=dict(showgrid=True),
        height=600,
        margin=dict(l=50, r=50, t=80, b=50),
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        )
    )
    
    return dcc.Graph(figure=fig)

def make_files_table():
    """Create a table showing all processed files"""
    if not all_processed_files:
        return html.Div("No files were processed")
    
    # Create a DataFrame for the table
    files_df = pd.DataFrame(all_processed_files)
    
    # Create the table using dash_table
    table = dash_table.DataTable(
        data=files_df.to_dict('records'),
        columns=[
            {"name": "Source", "id": "source"},
            {"name": "File Name", "id": "filename"},
            {"name": "DMT/Entity Type", "id": "dmt_type"},
            {"name": "Lot Number", "id": "lot_number"},
            {"name": "File Date/Time", "id": "file_datetime"},
            {"name": "Full Path", "id": "full_path"}
        ],
        style_table={'overflowX': 'auto'},
        style_cell={
            'textAlign': 'left',
            'padding': '10px',
            'fontFamily': 'Arial'
        },
        style_header={
            'backgroundColor': 'rgb(230, 230, 230)',
            'fontWeight': 'bold'
        },
        style_data_conditional=[
            {
                'if': {'filter_query': '{source} = TEST'},
                'backgroundColor': 'rgba(173, 216, 230, 0.3)',
            },
            {
                'if': {'filter_query': '{source} = POR'},
                'backgroundColor': 'rgba(255, 182, 193, 0.3)',
            }
        ],
        page_size=20,
        sort_action="native",
        filter_action="native"
    )
    
    return html.Div([
        html.H3(f"Processed Files ({len(all_processed_files)} total)"),
        table
    ])

# ===== EXCEL EXPORT FUNCTIONS =====

def create_test_export_excel(offset=0):
    """Create Excel export data for TEST statistics by COATER and File Name with offset applied"""
    if test_df.empty:
        return None
    
    # Apply offset to TEST data
    modified_test_df = apply_offset_to_test_data(test_df, offset)
    
    # Filter for thickness data only
    thickness_data = modified_test_df[modified_test_df['Label'] == 'Layer 1 Thickness'].copy()
    
    if thickness_data.empty:
        return None
    
    # Get unique file names sorted by datetime
    files_with_time = thickness_data.groupby('FileName')['datetime'].first().sort_values()
    file_names = files_with_time.index.tolist()
    
    # Initialize result DataFrame
    coater_order = ['1301', '1401', '1303', '1403', '1302', '1402', '1304', '1404']
    
    # Create column names: 3 columns per COATER
    columns = ['File Name']
    for coater in coater_order:
        columns.extend([
            f'{coater} Mean Thickness (Å)',
            f'{coater} Avg delta to previous run',
            f'{coater} std dev'
        ])
    
    export_data = []
    previous_means = {}  # Store previous run means for delta calculation
    
    for file_idx, filename in enumerate(file_names):
        row_data = {'File Name': filename}
        
        # Get data for this file
        file_data = thickness_data[thickness_data['FileName'] == filename]
        
        for coater in coater_order:
            coater_data = file_data[file_data['COATER'] == coater]['Datum']
            
            if len(coater_data) > 0:
                mean_thickness = round(coater_data.mean(), 1)
                std_dev = round(coater_data.std(), 1)
                
                # Calculate delta from previous run (use absolute value)
                if file_idx == 0:  # First file, no previous run
                    delta_to_previous = ''
                else:
                    if coater in previous_means:
                        # Use absolute value for delta
                        delta_to_previous = round(abs(mean_thickness - previous_means[coater]), 1)
                    else:
                        delta_to_previous = ''
                
                # Store current mean for next iteration
                previous_means[coater] = mean_thickness
                
                # Add to row
                row_data[f'{coater} Mean Thickness (Å)'] = mean_thickness
                row_data[f'{coater} Avg delta to previous run'] = delta_to_previous
                row_data[f'{coater} std dev'] = std_dev
            else:
                # No data for this coater in this file
                row_data[f'{coater} Mean Thickness (Å)'] = ''
                row_data[f'{coater} Avg delta to previous run'] = ''
                row_data[f'{coater} std dev'] = ''
        
        export_data.append(row_data)
    
    # Create DataFrame
    export_df = pd.DataFrame(export_data, columns=columns)
    
    # Create Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, sheet_name='TEST Statistical Summary', index=False)
        
        # Get workbook and worksheet for formatting
        workbook = writer.book
        worksheet = writer.sheets['TEST Statistical Summary']
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 40)  # Cap at 40
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
    
    output.seek(0)
    return output.getvalue()

def create_test_summary_table_data(offset=0):
    """Create table data for TEST statistics by COATER and File Name matching Excel export format"""
    if test_df.empty:
        return []
    
    # Apply offset to TEST data
    modified_test_df = apply_offset_to_test_data(test_df, offset)
    
    # Filter for thickness data only
    thickness_data = modified_test_df[modified_test_df['Label'] == 'Layer 1 Thickness'].copy()
    
    if thickness_data.empty:
        return []
    
    # Get unique file names sorted by datetime
    files_with_time = thickness_data.groupby('FileName')['datetime'].first().sort_values()
    file_names = files_with_time.index.tolist()
    
    # Initialize result data
    coater_order = ['1301', '1401', '1303', '1403', '1302', '1402', '1304', '1404']
    
    table_data = []
    previous_means = {}  # Store previous run means for delta calculation
    
    for file_idx, filename in enumerate(file_names):
        row_data = {'File Name': filename}
        
        # Get data for this file
        file_data = thickness_data[thickness_data['FileName'] == filename]
        
        for coater in coater_order:
            coater_data = file_data[file_data['COATER'] == coater]['Datum']
            
            if len(coater_data) > 0:
                mean_thickness = round(coater_data.mean(), 1)
                std_dev = round(coater_data.std(), 1)
                
                # Calculate delta from previous run (use absolute value)
                if file_idx == 0:  # First file, no previous run
                    delta_to_previous = ''
                else:
                    if coater in previous_means:
                        # Use absolute value for delta
                        delta_to_previous = round(abs(mean_thickness - previous_means[coater]), 1)
                    else:
                        delta_to_previous = ''
                
                # Store current mean for next iteration
                previous_means[coater] = mean_thickness
                
                # Add to row
                row_data[f'{coater} Mean Thickness (Å)'] = mean_thickness
                row_data[f'{coater} Avg delta to previous run'] = delta_to_previous
                row_data[f'{coater} std dev'] = std_dev
            else:
                # No data for this coater in this file
                row_data[f'{coater} Mean Thickness (Å)'] = ''
                row_data[f'{coater} Avg delta to previous run'] = ''
                row_data[f'{coater} std dev'] = ''
        
        table_data.append(row_data)
    
    return table_data

def create_test_summary_table(offset=0):
    """Create dash table for TEST statistics by COATER and File Name matching Excel export format"""
    table_data = create_test_summary_table_data(offset)
    
    if not table_data:
        return html.Div("No TEST data available for summary table")
    
    # Create column definitions
    coater_order = ['1301', '1401', '1303', '1403', '1302', '1402', '1304', '1404']
    
    columns = [{'name': 'File Name', 'id': 'File Name', 'type': 'text'}]
    
    for coater in coater_order:
        columns.extend([
            {'name': f'{coater} Mean (Å)', 'id': f'{coater} Mean Thickness (Å)', 'type': 'numeric'},
            {'name': f'{coater} Δ Prev', 'id': f'{coater} Avg delta to previous run', 'type': 'numeric'},
            {'name': f'{coater} StdDev', 'id': f'{coater} std dev', 'type': 'numeric'}
        ])
    
    # Create the dash table
    table = dash_table.DataTable(
        data=table_data,
        columns=columns,
        style_table={'overflowX': 'auto', 'minWidth': '100%'},
        style_cell={
            'textAlign': 'center',
            'padding': '8px',
            'fontFamily': 'Arial',
            'fontSize': '11px',
            'minWidth': '80px',
            'maxWidth': '120px'
        },
        style_header={
            'backgroundColor': 'rgb(230, 230, 230)',
            'fontWeight': 'bold',
            'fontSize': '10px',
            'textAlign': 'center'
        },
        style_data_conditional=[
            {
                'if': {'column_id': 'File Name'},
                'textAlign': 'left',
                'fontWeight': 'bold',
                'backgroundColor': 'rgb(248, 248, 255)',
                'minWidth': '200px'
            }
        ] + [
            {
                'if': {'column_id': f'{coater} Mean Thickness (Å)'},
                'backgroundColor': '#e8f4f8'
            } for coater in coater_order
        ] + [
            {
                'if': {'column_id': f'{coater} Avg delta to previous run'},
                'backgroundColor': '#fff8e8'
            } for coater in coater_order
        ] + [
            {
                'if': {'column_id': f'{coater} std dev'},
                'backgroundColor': '#f8fff8'
            } for coater in coater_order
        ],
        page_size=20,
        sort_action="native",
        filter_action="native"
    )
    
    # Add offset info to title
    title_text = f"TEST Statistical Summary by File and COATER (Offset: {offset:+.1f} Å)" if offset != 0 else "TEST Statistical Summary by File and COATER"
    
    return html.Div([
        html.H3(title_text),
        html.P([
            html.Strong("Mean: "), "Average thickness per coater per file | ",
            html.Strong("Δ Prev: "), "Absolute delta from previous run | ",
            html.Strong("StdDev: "), "Standard deviation within run"
        ], style={'fontSize': '12px', 'marginBottom': '15px', 'fontStyle': 'italic'}),
        table
    ])

# ===== DASH APP SETUP =====

app = dash.Dash(__name__)

app.layout = html.Div([
    html.H1("Comprehensive Thickness Data Analysis: TEST vs POR", 
            style={'textAlign': 'center', 'marginBottom': '30px'}),
    
    # Matching Offset Controls
    html.Div([
        html.H3("Matching Offset Control"),
        html.Div([
            html.Label("Matching Offset (Angstrom): ", style={'marginRight': '10px'}),
            dcc.Dropdown(
                id='offset-dropdown',
                options=[{'label': f'{i/10:.1f}', 'value': i/10} for i in range(-500, 501, 5)],  # -50.0 to +50.0 in 0.5 increments
                value=0,
                style={'width': '150px', 'display': 'inline-block'}
            ),
            html.Span(" (Applied to TEST data only)", style={'marginLeft': '10px', 'fontStyle': 'italic'})
        ], style={'display': 'flex', 'alignItems': 'center'})
    ], style={'margin': '20px 0', 'padding': '15px', 'backgroundColor': '#e8f4f8', 'borderRadius': '5px', 'border': '2px solid #007acc'}),
    
    # Summary Statistics
    html.Div([
        html.H2("Data Summary"),
        html.Div([
            html.Div([
                html.H3(f"TEST Data: {len(test_df[test_df['Label'] == 'Layer 1 Thickness'])} thickness measurements", 
                        style={'color': 'blue'}),
                html.P(f"Source: XML files from TEST folder (Entities: {', '.join(test_df['Entity'].unique()) if not test_df.empty else 'None'})")
            ], style={'width': '48%', 'display': 'inline-block', 'verticalAlign': 'top'}),
            
            html.Div([
                html.H3(f"POR Data: {len(por_df[por_df['Label'] == 'Layer 1 Thickness'])} thickness measurements", 
                        style={'color': 'red'}),
                html.P(f"Source: CSV file from POR folder (Entities: {', '.join(por_df['Entity'].unique()) if not por_df.empty else 'None'})")
            ], style={'width': '48%', 'display': 'inline-block', 'verticalAlign': 'top', 'marginLeft': '4%'})
        ])
    ], style={'margin': '20px 0', 'padding': '15px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'}),
    
    # Statistical Comparison Table
    html.Div(id='stats-table'),
    
    html.Hr(),
    
    # Entity Comparison Plots
    html.H2("Entity and COATER Analysis"),
    html.Div(id='entity-comparison'),
    
    html.Hr(),
    
    # Raw Data Comparison
    html.H2("Raw Thickness Data Comparison"),
    html.Div(id='boxplot-comparison'),
    
    html.Hr(),
    
    # TEST Data Export
    html.H2("TEST Data Export"),
    html.Div([
        html.P("Export TEST data statistical summary to Excel format with metrics by COATER and File Name."),
        html.Button(
            "Export TEST Data to Excel",
            id="export-button",
            n_clicks=0,
            style={
                'backgroundColor': '#007acc',
                'color': 'white',
                'padding': '10px 20px',
                'border': 'none',
                'borderRadius': '5px',
                'cursor': 'pointer',
                'fontSize': '14px',
                'marginBottom': '15px'
            }
        ),
        dcc.Download(id="download-dataframe-xlsx")
    ], style={'margin': '15px 0', 'padding': '15px', 'backgroundColor': '#f0f8ff', 'borderRadius': '5px'}),
    
    html.Hr(),
    
    # TEST Data Summary Table (matching Excel export format)
    html.H2("TEST Data Statistical Summary Table"),
    html.P([
        "Statistical summary showing mean thickness, run-to-run deltas (absolute values), and standard deviations by COATER for each file. ",
        "This table matches the Excel export format."
    ], style={'fontSize': '14px', 'marginBottom': '15px'}),
    html.Div(id='test-summary-table'),
    
    html.Hr(),
    
    # TEST Time Series Analysis
    html.H2("TEST Data Time Series"),
    html.Div(id='test-time-series'),
    
    html.Hr(),
    
    # TEST DECK Time Series Analysis
    html.H2("TEST Data Time Series by DECK"),
    html.P([
        "Grouped by DECK: ",
        html.Strong("13-L"), " (COATER 1301, 1302) | ",
        html.Strong("13-R"), " (COATER 1303, 1304) | ",
        html.Strong("14-L"), " (COATER 1401, 1402) | ",
        html.Strong("14-R"), " (COATER 1403, 1404)"
    ], style={'fontSize': '12px', 'fontStyle': 'italic', 'marginBottom': '15px'}),
    html.Div(id='test-deck-time-series'),
    
    html.Hr(),
    
    # TEST Radius Spline Analysis
    html.H2("TEST Thickness vs Radius by DECK (Spline Analysis)"),
    html.P([
        "Interactive spline analysis showing thickness uniformity across wafer radius (0-150mm). ",
        "Select specific files to analyze or view all data combined."
    ], style={'fontSize': '14px', 'marginBottom': '15px'}),
    
    # File selection for spline plot
    html.Div([
        html.Label("Select Files for Spline Analysis:", style={'fontWeight': 'bold', 'marginBottom': '5px'}),
        dcc.Dropdown(
            id='spline-file-dropdown',
            options=[{'label': file, 'value': file} for file in sorted(test_df['FileName'].unique())] if not test_df.empty else [],
            value=None,  # Start with all files
            multi=True,
            placeholder="Select files (leave empty for all files)",
            style={'marginBottom': '15px'}
        )
    ], style={'margin': '15px 0'}),
    
    html.Div(id='test-radius-spline'),
    
    html.Hr(),
    
    # POR Radius Spline Analysis
    html.H2("POR Thickness vs Radius by COATER (Spline Analysis)"),
    html.P([
        "Interactive spline analysis showing POR thickness uniformity across wafer radius (0-150mm). ",
        "Select specific entities to analyze or view all data combined."
    ], style={'fontSize': '14px', 'marginBottom': '15px'}),
    
    # Entity selection for spline plot
    html.Div([
        html.Label("Select Entities for Spline Analysis:", style={'fontWeight': 'bold', 'marginBottom': '5px'}),
        dcc.Dropdown(
            id='por-entity-dropdown',
            options=[{'label': entity, 'value': entity} for entity in sorted(por_df['Entity'].unique())] if not por_df.empty else [],
            value=None,  # Start with all entities
            multi=True,
            placeholder="Select entities (leave empty for all entities)",
            style={'marginBottom': '15px'}
        )
    ], style={'margin': '15px 0'}),
    
    html.Div(id='por-radius-spline'),
    
    html.Hr(),
    
    # Processed Files Information
    html.H2("Data Source Files"),
    make_files_table()
])

# Callback to update all plots when offset changes
@app.callback(
    [Output('stats-table', 'children'),
     Output('entity-comparison', 'children'),
     Output('boxplot-comparison', 'children'),
     Output('test-summary-table', 'children'),
     Output('test-time-series', 'children'),
     Output('test-deck-time-series', 'children'),
     Output('test-radius-spline', 'children'),
     Output('por-radius-spline', 'children')],
    [Input('offset-dropdown', 'value'),
     Input('spline-file-dropdown', 'value'),
     Input('por-entity-dropdown', 'value')]
)
def update_plots(offset, selected_files, selected_entities):
    """Update all plots and tables when the matching offset changes, files, or entities are selected"""
    return (
        make_statistics_comparison_table(offset),
        create_entity_comparison_plots(),
        make_source_comparison_boxplot(offset),
        create_test_summary_table(offset),
        make_test_time_series_plot(offset),
        make_test_deck_time_series_plot(offset),
        make_test_radius_spline_plot(selected_files, offset),
        make_por_radius_spline_plot(selected_entities)
    )

# Callback for Excel export
@app.callback(
    Output("download-dataframe-xlsx", "data"),
    [Input("export-button", "n_clicks"),
     Input('offset-dropdown', 'value')],
    prevent_initial_call=True,
)
def generate_excel_export(n_clicks, offset):
    """Generate and download Excel file with TEST data statistics including offset adjustment"""
    if n_clicks > 0:
        excel_data = create_test_export_excel(offset)
        if excel_data:
            offset_info = f"_offset{offset:+.1f}" if offset != 0 else ""
            return dcc.send_bytes(
                excel_data,
                f"TEST_Statistical_Summary{offset_info}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
    return None

if __name__ == '__main__':
    print(f"\nData Loading Summary:")
    print(f"TEST DataFrame: {len(test_df)} records")
    print(f"POR DataFrame: {len(por_df)} records")
    print(f"Combined DataFrame: {len(combined_df)} records")
    
    if not test_df.empty:
        print(f"TEST Columns: {test_df.columns.tolist()}")
        test_thickness = test_df[test_df['Label'] == 'Layer 1 Thickness']
        print(f"TEST Thickness data points: {len(test_thickness)}")
    
    if not por_df.empty:
        print(f"POR Columns: {por_df.columns.tolist()}")
        por_thickness = por_df[por_df['Label'] == 'Layer 1 Thickness']
        print(f"POR Thickness data points: {len(por_thickness)}")
    
    print(f"\nStarting Dash app on http://127.0.0.1:8050/")
    app.run(debug=True)