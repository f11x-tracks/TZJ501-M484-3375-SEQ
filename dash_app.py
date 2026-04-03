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
from scipy.interpolate import griddata

import xml.etree.ElementTree as ET

import plotly.express as px
import plotly.graph_objects as go

# ===== COATER CONFIGURATION =====
# TEST Data: 8 COATER positions based on wafer order within each XML file (easily configurable)
# Each XML file starts fresh with wafer #1
TEST_COATER_BASE_PATTERN = ['1301', '1401', '1303', '1403', '1302', '1402', '1304', '1404']

# TEST Entity: Hardcoded entity for TEST data (easily configurable)
TEST_ENTITY = 'TZJ501'  # Change this value to update the TEST entity

# DECK Mapping: Group coaters into decks (easily configurable)
COATER_TO_DECK_MAPPING = {
    '1301': '13-L', '1302': '13-L',
    '1303': '13-R', '1304': '13-R', 
    '1401': '14-L', '1402': '14-L',
    '1403': '14-R', '1404': '14-R'
}

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
                        'Entity': TEST_ENTITY  # Use configurable TEST entity
                    })
                except (TypeError, ValueError):
                    continue
        
        # Add to processed files list if we got here without errors
        processed_files.append({
            'filename': os.path.basename(file),
            'full_path': file,
            'dmt_type': dmt,
            'lot_number': lot_number,
            'file_datetime': file_time.strftime('%Y-%m-%d %H:%M:%S')
        })
        
    except Exception as e:
        print(f"Error processing file {file}: {e}")
        continue

# Read condition data
condition_file = 'acondition.txt'
try:
    # Read the file and filter out comment lines (starting with #)
    with open(condition_file, 'r') as f:
        lines = f.readlines()
    
    # Filter out lines starting with # and empty lines
    filtered_lines = [line for line in lines if line.strip() and not line.strip().startswith('#')]
    
    # Create a temporary string with filtered content
    from io import StringIO
    filtered_content = ''.join(filtered_lines)
    
    # Read CSV from the filtered content - handle both space and tab delimited
    # Use regex pattern to handle any combination of tabs and spaces as delimiters
    import re
    
    if filtered_lines:
        # Use regex to handle any whitespace (tabs, spaces, or combinations) as delimiters
        # This will work regardless of whether the file uses tabs, spaces, or mixed delimiters
        conditions_df = pd.read_csv(StringIO(filtered_content), sep=r'\s+', engine='python')
    else:
        # Empty file, create empty dataframe
        conditions_df = pd.DataFrame(columns=['WaferID', 'Condition', 'Condition2'])
    
    # Clean up column names (remove any extra whitespace)
    conditions_df.columns = conditions_df.columns.str.strip()
    conditions_df['WaferID'] = conditions_df['WaferID'].str.strip()
    conditions_df['Condition'] = conditions_df['Condition'].str.strip()
    
    # Handle Condition2 column if it exists
    if 'Condition2' in conditions_df.columns:
        conditions_df['Condition2'] = conditions_df['Condition2'].str.strip().fillna('')
        # Combine Condition and Condition2 with parentheses format
        # Only add parentheses if Condition2 is not empty
        conditions_df['CombinedCondition'] = conditions_df.apply(
            lambda row: f"{row['Condition']} ({row['Condition2']})" if row['Condition2'] 
            else row['Condition'], axis=1)
    else:
        # If no Condition2 column, use original Condition
        conditions_df['CombinedCondition'] = conditions_df['Condition']
    
    print(f"Loaded {len(conditions_df)} condition records from {condition_file}")
except Exception as e:
    print(f"Could not read condition file {condition_file}: {e}")
    conditions_df = pd.DataFrame(columns=['WaferID', 'Condition'])

df = pd.DataFrame(records)

# Merge condition data with main dataframe
if not conditions_df.empty and 'CombinedCondition' in conditions_df.columns:
    # Use CombinedCondition as the main Condition for display
    df = df.merge(conditions_df[['WaferID', 'CombinedCondition']], on='WaferID', how='left')
    # Rename CombinedCondition to Condition for consistency with existing code
    df = df.rename(columns={'CombinedCondition': 'Condition'})
else:
    # Create empty Condition column if no condition data available
    df['Condition'] = 'NA'

# Fill missing conditions with 'NA'
df['Condition'] = df['Condition'].fillna('NA')

if not conditions_df.empty:
    print(f"Added condition data. Conditions found: {df['Condition'].value_counts().to_dict()}")
else:
    print("No condition data available - all conditions set to 'NA'")

# Dash app
app = dash.Dash(__name__)

def make_boxplot(label, y_range=None, filtered_df=None):
    working_df = filtered_df if filtered_df is not None else df
    dff = working_df[working_df['Label'] == label]
    if dff.empty:
        return html.Div(f"No data for {label}")
    fig = px.box(
        dff,
        x=dff['datetime'].dt.strftime('%Y-%m-%d %H:%M:%S'),
        y='Datum',
        color='dmt',
        points='all',
        title=f'Boxplot of {label} over Time',
        labels={'Datum': label, 'datetime': 'File DateTime', 'dmt': 'DMT Type'}
    )
    fig.update_layout(xaxis_title='File DateTime', yaxis_title=label)
    if y_range and label == 'Layer 1 Thickness':
        fig.update_layout(yaxis=dict(range=y_range))
    return dcc.Graph(figure=fig)

def make_wafer_plots(label):
    dff = df[df['Label'] == label]
    if dff.empty:
        return html.Div(f"No data for {label}")
    
    unique_wafers = sorted(dff['WaferID'].unique())
    plots = []
    
    for wafer_id in unique_wafers:
        wafer_data = dff[dff['WaferID'] == wafer_id]
        if not wafer_data.empty:
            fig = px.box(
                wafer_data,
                x=wafer_data['datetime'].dt.strftime('%Y-%m-%d %H:%M:%S'),
                y='Datum',
                color='dmt',
                points='all',
                title=f'{label} - WaferID: {wafer_id}',
                labels={'Datum': label, 'datetime': 'File DateTime', 'dmt': 'DMT Type'}
            )
            fig.update_layout(
                xaxis_title='File DateTime', 
                yaxis_title=label,
                height=400,
                margin=dict(l=50, r=50, t=50, b=50)
            )
            plots.append(dcc.Graph(figure=fig))
    
    return html.Div(plots)

def make_scatter_plot():
    # Filter data that has location information for proper pairing
    df_with_location = df[df['location_id'].notna()].copy()
    
    if df_with_location.empty:
        return html.Div("No location data available for scatter plot")
    
    # Separate GoF and Thickness data
    gof_data = df_with_location[df_with_location['Label'] == 'Goodness-of-Fit'][['datetime', 'WaferID', 'dmt', 'LotNumber', 'location_id', 'Datum']].rename(columns={'Datum': 'GoodnessOfFit'})
    thickness_data = df_with_location[df_with_location['Label'] == 'Layer 1 Thickness'][['datetime', 'WaferID', 'dmt', 'LotNumber', 'location_id', 'Datum']].rename(columns={'Datum': 'Layer1Thickness'})
    
    # Merge based on location (same measurement point)
    merged_data = pd.merge(gof_data, thickness_data, on=['datetime', 'WaferID', 'dmt', 'LotNumber', 'location_id'], how='inner')
    
    if merged_data.empty:
        return html.Div("No paired measurement data available for scatter plot")
    
    fig = px.scatter(
        merged_data,
        x='GoodnessOfFit',
        y='Layer1Thickness',
        color='dmt',
        symbol='WaferID',
        title='Layer 1 Thickness vs Goodness-of-Fit (Same Measurement Points)',
        labels={
            'GoodnessOfFit': 'Goodness-of-Fit',
            'Layer1Thickness': 'Layer 1 Thickness',
            'dmt': 'DMT Type'
        },
        hover_data=['WaferID', 'LotNumber', 'datetime']
    )
    
    fig.update_layout(
        xaxis_title='Goodness-of-Fit',
        yaxis_title='Layer 1 Thickness',
        height=600,
        margin=dict(l=50, r=50, t=50, b=50)
    )
    
    return dcc.Graph(figure=fig)

def make_radius_thickness_plots(y_range=None, show_trend_legend=True, filtered_df=None):
    working_df = filtered_df if filtered_df is not None else df
    # Filter for Layer 1 Thickness data with valid RADIUS
    thickness_data = working_df[(working_df['Label'] == 'Layer 1 Thickness') & (working_df['RADIUS'].notna())].copy()
    
    if thickness_data.empty:
        return html.Div("No Layer 1 Thickness data with RADIUS available")
    
    # Create single scatter plot
    fig = go.Figure()
    
    # Get unique wafer IDs and assign colors/symbols
    unique_wafers = sorted(thickness_data['WaferID'].unique())
    
    # Define colors for each wafer (expanded to 12 colors)
    plotly_colors = [
        '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
        '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
        '#aec7e8', '#ffbb78'
    ]
    
    # Add scatter points colored by WaferID
    for i, wafer_id in enumerate(unique_wafers):
        wafer_data = thickness_data[thickness_data['WaferID'] == wafer_id].copy()
        
        if wafer_data.empty:
            continue
        
        # Assign specific color for this wafer
        wafer_color = plotly_colors[i % len(plotly_colors)]
            
        # Add scatter points
        scatter_trace = go.Scatter(
            x=wafer_data['RADIUS'],
            y=wafer_data['Datum'],
            mode='markers',
            name=f'WaferID: {wafer_id}',
            marker=dict(size=8, color=wafer_color),
            text=wafer_data['datetime'].dt.strftime('%Y-%m-%d %H:%M:%S'),
            customdata=wafer_data[['dmt', 'LotNumber']],
            hovertemplate='<b>%{fullData.name}</b><br>' +
                          'RADIUS: %{x:.2f}<br>' +
                          'Thickness: %{y:.2f}<br>' +
                          'DMT Type: %{customdata[0]}<br>' +
                          'Lot Number: %{customdata[1]}<br>' +
                          'DateTime: %{text}<br>' +
                          '<extra></extra>'
        )
        fig.add_trace(scatter_trace)
        
        # Add LOWESS trend line for each wafer if there are enough points
        if len(wafer_data) >= 5:  # Need at least 5 points for LOWESS
            try:
                # Sort by RADIUS for proper trend line
                wafer_sorted = wafer_data.sort_values('RADIUS')
                
                # Apply LOWESS smoothing
                # frac parameter controls smoothness (0.1-1.0, smaller = less smooth)
                lowess_result = lowess(wafer_sorted['Datum'], wafer_sorted['RADIUS'], 
                                     frac=0.3, it=3, return_sorted=True)
                
                # Extract smoothed values
                x_smooth = lowess_result[:, 0]
                y_smooth = lowess_result[:, 1]
                
                # Add LOWESS trend line with matching color
                fig.add_trace(go.Scatter(
                    x=x_smooth,
                    y=y_smooth,
                    mode='lines',
                    name=f'LOWESS: {wafer_id}',
                    line=dict(dash='dash', width=2, color=wafer_color),
                    showlegend=show_trend_legend,   # Controlled by parameter
                    hoverinfo='skip'   # Don't show hover for trend line
                ))
                
            except Exception as e:
                # Fallback to simple linear regression if LOWESS fails
                try:
                    coeffs = np.polyfit(wafer_sorted['RADIUS'], wafer_sorted['Datum'], 1)
                    x_fit = np.linspace(wafer_sorted['RADIUS'].min(), wafer_sorted['RADIUS'].max(), 50)
                    y_fit = np.polyval(coeffs, x_fit)
                    
                    fig.add_trace(go.Scatter(
                        x=x_fit,
                        y=y_fit,
                        mode='lines',
                        name=f'Linear: {wafer_id}',
                        line=dict(dash='dot', width=2, color=wafer_color),
                        showlegend=show_trend_legend,
                        hoverinfo='skip'
                    ))
                except:
                    pass  # Skip trend line if all methods fail
    
    # Update layout
    fig.update_layout(
        title='Layer 1 Thickness vs RADIUS by WaferID (with LOWESS Trend Lines)',
        xaxis_title='RADIUS',
        yaxis_title='Layer 1 Thickness',
        xaxis=dict(range=[0, 150], dtick=15),
        height=600,
        margin=dict(l=50, r=50, t=50, b=50),
        showlegend=True
    )
    
    # Apply y-axis range if provided
    if y_range:
        fig.update_layout(yaxis=dict(range=y_range))
    
    return dcc.Graph(figure=fig)

def make_radius_thickness_by_condition_plots(y_range=None, show_trend_legend=True, filtered_df=None):
    working_df = filtered_df if filtered_df is not None else df
    # Filter for Layer 1 Thickness data with valid RADIUS and non-NA conditions
    thickness_data = working_df[(working_df['Label'] == 'Layer 1 Thickness') & 
                       (working_df['RADIUS'].notna()) & 
                       (working_df['Condition'] != 'NA')].copy()
    
    if thickness_data.empty:
        return html.Div("No Layer 1 Thickness data with RADIUS and conditions available")
    
    # Create single scatter plot
    fig = go.Figure()
    
    # Get unique conditions
    unique_conditions = sorted(thickness_data['Condition'].unique())
    
    # Define colors for each condition (expanded to 12 colors)
    plotly_colors = [
        '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
        '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
        '#aec7e8', '#ffbb78'
    ]
    
    # Add scatter points colored by Condition
    for i, condition in enumerate(unique_conditions):
        condition_data = thickness_data[thickness_data['Condition'] == condition].copy()
        
        if condition_data.empty:
            continue
        
        # Assign specific color for this condition
        condition_color = plotly_colors[i % len(plotly_colors)]
            
        # Add scatter points
        scatter_trace = go.Scatter(
            x=condition_data['RADIUS'],
            y=condition_data['Datum'],
            mode='markers',
            name=f'Condition: {condition}',
            marker=dict(size=8, color=condition_color),
            text=condition_data['datetime'].dt.strftime('%Y-%m-%d %H:%M:%S'),
            customdata=condition_data[['dmt', 'WaferID', 'LotNumber']],
            hovertemplate='<b>%{fullData.name}</b><br>' +
                          'RADIUS: %{x:.2f}<br>' +
                          'Thickness: %{y:.2f}<br>' +
                          'WaferID: %{customdata[1]}<br>' +
                          'DMT Type: %{customdata[0]}<br>' +
                          'Lot Number: %{customdata[2]}<br>' +
                          'DateTime: %{text}<br>' +
                          '<extra></extra>'
        )
        fig.add_trace(scatter_trace)
        
        # Add LOWESS trend line for each condition if there are enough points
        if len(condition_data) >= 5:  # Need at least 5 points for LOWESS
            try:
                # Sort by RADIUS for proper trend line
                condition_sorted = condition_data.sort_values('RADIUS')
                
                # Apply LOWESS smoothing
                lowess_result = lowess(condition_sorted['Datum'], condition_sorted['RADIUS'], 
                                     frac=0.3, it=3, return_sorted=True)
                
                # Extract smoothed values
                x_smooth = lowess_result[:, 0]
                y_smooth = lowess_result[:, 1]
                
                # Add LOWESS trend line with matching color
                fig.add_trace(go.Scatter(
                    x=x_smooth,
                    y=y_smooth,
                    mode='lines',
                    name=f'LOWESS: {condition}',
                    line=dict(dash='dash', width=3, color=condition_color),
                    showlegend=show_trend_legend,   # Controlled by parameter
                    hoverinfo='skip'   # Don't show hover for trend line
                ))
                
            except Exception as e:
                # Fallback to simple linear regression if LOWESS fails
                try:
                    coeffs = np.polyfit(condition_sorted['RADIUS'], condition_sorted['Datum'], 1)
                    x_fit = np.linspace(condition_sorted['RADIUS'].min(), condition_sorted['RADIUS'].max(), 50)
                    y_fit = np.polyval(coeffs, x_fit)
                    
                    fig.add_trace(go.Scatter(
                        x=x_fit,
                        y=y_fit,
                        mode='lines',
                        name=f'Linear: {condition}',
                        line=dict(dash='dot', width=3, color=condition_color),
                        showlegend=show_trend_legend,
                        hoverinfo='skip'
                    ))
                except:
                    pass  # Skip trend line if all methods fail
    
    # Update layout
    fig.update_layout(
        title='Layer 1 Thickness vs RADIUS by Condition (with LOWESS Trend Lines)',
        xaxis_title='RADIUS',
        yaxis_title='Layer 1 Thickness',
        xaxis=dict(range=[0, 150], dtick=15),
        height=600,
        margin=dict(l=50, r=50, t=50, b=50),
        showlegend=True
    )
    
    # Apply y-axis range if provided
    if y_range:
        fig.update_layout(yaxis=dict(range=y_range))
    
    return dcc.Graph(figure=fig)

def make_wafer_contour_plot(wafer_id, filtered_df=None):
    """Create a contour plot for Layer 1 Thickness using x_wafer_loc and y_wafer_loc for a specific wafer"""
    working_df = filtered_df if filtered_df is not None else df
    
    if wafer_id is None:
        return html.Div("Please select a wafer to display contour plot")
    
    # Filter for the selected wafer and Layer 1 Thickness data
    wafer_thickness_data = working_df[
        (working_df['WaferID'] == wafer_id) & 
        (working_df['Label'] == 'Layer 1 Thickness') &
        (working_df['XWaferLoc'].notna()) &
        (working_df['YWaferLoc'].notna())
    ].copy()
    
    if wafer_thickness_data.empty:
        return html.Div(f"No Layer 1 Thickness data with coordinates found for WaferID: {wafer_id}")
    
    # Convert coordinates to numeric
    try:
        wafer_thickness_data.loc[:, 'X'] = pd.to_numeric(wafer_thickness_data['XWaferLoc'])
        wafer_thickness_data.loc[:, 'Y'] = pd.to_numeric(wafer_thickness_data['YWaferLoc'])
        wafer_thickness_data.loc[:, 'Z'] = pd.to_numeric(wafer_thickness_data['Datum'])
    except (ValueError, TypeError):
        return html.Div(f"Error converting coordinate data to numeric for WaferID: {wafer_id}")
    
    # Remove any rows with invalid coordinates
    wafer_thickness_data = wafer_thickness_data.dropna(subset=['X', 'Y', 'Z'])
    
    if wafer_thickness_data.empty:
        return html.Div(f"No valid coordinate data found for WaferID: {wafer_id}")
    
    # Get unique coordinates and values
    x = wafer_thickness_data['X'].values
    y = wafer_thickness_data['Y'].values
    z = wafer_thickness_data['Z'].values
    
    # Create a grid for interpolation
    xi = np.linspace(-150, 150, 100)
    yi = np.linspace(-150, 150, 100)
    xi_grid, yi_grid = np.meshgrid(xi, yi)
    
    try:
        # Interpolate data onto grid using griddata
        zi = griddata((x, y), z, (xi_grid, yi_grid), method='linear')
        
        # Create the contour plot
        fig = go.Figure()
        
        # Add contour plot
        contour = go.Contour(
            x=xi,
            y=yi,
            z=zi,
            colorscale='Viridis',
            colorbar=dict(title="Layer 1 Thickness"),
            contours=dict(
                start=np.nanmin(zi),
                end=np.nanmax(zi),
                size=(np.nanmax(zi) - np.nanmin(zi)) / 20
            )
        )
        fig.add_trace(contour)
        
        # Add scatter points to show actual measurement locations
        scatter = go.Scatter(
            x=x,
            y=y,
            mode='markers',
            marker=dict(
                size=6,
                color=z,
                colorscale='Viridis',
                line=dict(width=1, color='black'),
                showscale=False
            ),
            text=[f'X: {xi:.1f}<br>Y: {yi:.1f}<br>Thickness: {zi:.2f}' 
                  for xi, yi, zi in zip(x, y, z)],
            hovertemplate='%{text}<extra></extra>',
            name='Measurement Points'
        )
        fig.add_trace(scatter)
        
        # Update layout
        fig.update_layout(
            title=f'Layer 1 Thickness Contour Plot - WaferID: {wafer_id}',
            xaxis_title='X Wafer Location (mm)',
            yaxis_title='Y Wafer Location (mm)',
            xaxis=dict(
                range=[-150, 150],
                dtick=30,
                scaleanchor="y",
                scaleratio=1
            ),
            yaxis=dict(
                range=[-150, 150],
                dtick=30
            ),
            height=600,
            width=600,
            margin=dict(l=50, r=50, t=50, b=50)
        )
        
        # Add circle outlines to show wafer boundary (assuming 150mm radius wafer)
        circle_angles = np.linspace(0, 2*np.pi, 100)
        circle_x = 150 * np.cos(circle_angles)
        circle_y = 150 * np.sin(circle_angles)
        
        fig.add_trace(go.Scatter(
            x=circle_x,
            y=circle_y,
            mode='lines',
            line=dict(color='black', width=2, dash='dash'),
            name='Wafer Boundary (150mm)',
            showlegend=True,
            hoverinfo='skip'
        ))
        
        return dcc.Graph(figure=fig)
        
    except Exception as e:
        return html.Div(f"Error creating contour plot: {str(e)}")

def make_files_table():
    """Create a table showing all processed XML files"""
    if not processed_files:
        return html.Div("No files were processed")
    
    # Create a DataFrame for the table
    files_df = pd.DataFrame(processed_files)
    
    # Create the table using dash_table
    table = dash_table.DataTable(
        data=files_df.to_dict('records'),
        columns=[
            {"name": "File Name", "id": "filename"},
            {"name": "DMT Type", "id": "dmt_type"},
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
                'if': {'filter_query': '{dmt_type} = DMT102'},
                'backgroundColor': 'rgba(255, 182, 193, 0.3)',
            },
            {
                'if': {'filter_query': '{dmt_type} = DMT103'},
                'backgroundColor': 'rgba(173, 216, 230, 0.3)',
            }
        ],
        page_size=20,
        sort_action="native",
        filter_action="native"
    )
    
    return html.Div([
        html.H3(f"Processed XML Files ({len(processed_files)} total)"),
        table
    ])

def parse_combined_condition(combined_condition):
    """Parse combined condition like '1_Wfr1 (13-1)' into base condition and condition2"""
    if ' (' in combined_condition and combined_condition.endswith(')'):
        base_condition = combined_condition.split(' (')[0]
        condition2 = combined_condition.split(' (')[1][:-1]  # Remove the closing parenthesis
        return base_condition, condition2
    else:
        return combined_condition, ''

def make_statistical_summary_table(filtered_df=None):
    """Create a statistical summary table showing mean and std dev for each WaferID with radial groups"""
    working_df = filtered_df if filtered_df is not None else df
    if working_df.empty:
        return html.Div("No data available for statistical summary")
    
    # Get unique wafer IDs
    unique_wafers = sorted(working_df['WaferID'].unique())
    
    summary_data = []
    
    for wafer_id in unique_wafers:
        wafer_data = working_df[working_df['WaferID'] == wafer_id]
        
        # Get condition and slot for this wafer
        combined_condition = wafer_data['Condition'].iloc[0] if not wafer_data.empty else 'NA'
        base_condition, condition2 = parse_combined_condition(combined_condition)
        slot = wafer_data['Slot'].iloc[0] if not wafer_data.empty else 'Unknown'
        
        # Get datetime (file datetime) for this wafer
        file_datetime = wafer_data['datetime'].iloc[0] if not wafer_data.empty else None
        file_datetime_str = file_datetime.strftime('%Y-%m-%d %H:%M:%S') if file_datetime else 'Unknown'
        
        # Only process Layer 1 Thickness data
        thickness_data = wafer_data[wafer_data['Label'] == 'Layer 1 Thickness']
        label_data = thickness_data['Datum']
        
        if not label_data.empty:
            mean_val = label_data.mean()
            std_val = label_data.std()
            count_val = len(label_data)
            
            # Calculate standard deviations for radial groups
            # Center group: RADIUS 0-89
            center_data = thickness_data[(thickness_data['RADIUS'] >= 0) & (thickness_data['RADIUS'] < 89)]['Datum']
            center_std = round(center_data.std(), 4) if len(center_data) > 1 else None
            
            # Mid group: RADIUS 89-120
            mid_data = thickness_data[(thickness_data['RADIUS'] >= 89) & (thickness_data['RADIUS'] < 120)]['Datum']
            mid_std = round(mid_data.std(), 4) if len(mid_data) > 1 else None
            
            # Edge group: RADIUS 120-150
            edge_data = thickness_data[(thickness_data['RADIUS'] >= 120) & (thickness_data['RADIUS'] <= 150)]['Datum']
            edge_std = round(edge_data.std(), 4) if len(edge_data) > 1 else None
            
            summary_data.append({
                'Slot': slot,
                'WaferID': wafer_id,
                'File DateTime': file_datetime_str,
                'Condition': base_condition,
                'Condition2': condition2,
                'Measurement': 'Layer 1 Thickness',
                'Mean': round(mean_val, 1),
                'Std Dev': round(std_val, 1),
                'Count': count_val,
                'Center Std (0-89)': round(center_std, 1) if center_std is not None else None,
                'Mid Std (89-120)': round(mid_std, 1) if mid_std is not None else None,
                'Edge Std (120-150)': round(edge_std, 1) if edge_std is not None else None
            })
    
    if not summary_data:
        return html.Div("No statistical data to display")
    
    # Sort summary data by Slot first, then by Condition
    summary_data = sorted(summary_data, key=lambda x: (str(x['Slot']), x['Condition']))
    
    # Create the summary table using dash_table
    summary_table = dash_table.DataTable(
        data=summary_data,
        columns=[
            {"name": "Slot", "id": "Slot"},
            {"name": "Wafer ID", "id": "WaferID"},
            {"name": "File DateTime", "id": "File DateTime"},
            {"name": "Condition", "id": "Condition"},
            {"name": "Condition2", "id": "Condition2"},
            {"name": "Measurement Type", "id": "Measurement"},
            {"name": "Mean", "id": "Mean", "type": "numeric", "format": {"specifier": ".1f"}},
            {"name": "Std Dev", "id": "Std Dev", "type": "numeric", "format": {"specifier": ".1f"}},
            {"name": "Count", "id": "Count", "type": "numeric"},
            {"name": "Center Std (0-89)", "id": "Center Std (0-89)", "type": "numeric", "format": {"specifier": ".1f"}},
            {"name": "Mid Std (89-120)", "id": "Mid Std (89-120)", "type": "numeric", "format": {"specifier": ".1f"}},
            {"name": "Edge Std (120-150)", "id": "Edge Std (120-150)", "type": "numeric", "format": {"specifier": ".1f"}}
        ],
        style_table={'overflowX': 'auto'},
        style_cell={
            'textAlign': 'center',
            'padding': '8px',
            'fontFamily': 'Arial',
            'fontSize': '12px'
        },
        style_header={
            'backgroundColor': 'rgb(230, 230, 230)',
            'fontWeight': 'bold',
            'textAlign': 'center'
        },
        style_data_conditional=[
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': 'rgba(248, 248, 248, 0.8)',
            }
        ],
        sort_action="native",
        filter_action="native",
        page_size=30
    )
    
    return html.Div([
        html.H3(f"Statistical Summary by WaferID ({len(unique_wafers)} wafers)"),
        html.P("Mean, standard deviation, and radial group statistics for Layer 1 Thickness by wafer with condition"),
        html.P("Radial Groups: Center (0-89), Mid (89-120), Edge (120-150)", style={'fontStyle': 'italic', 'fontSize': '12px'}),
        summary_table
    ])

def export_full_data_to_excel():
    """Export the full dataframe with all data including RADIUS and Conditions to Excel"""
    try:
        # Generate timestamp for filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'DMT_Full_Data_{timestamp}.xlsx'
        
        if df.empty:
            print("No data to export")
            return None, False
        
        # Create a copy of the dataframe for export
        export_df = df.copy()
        
        # Format datetime column for better readability
        export_df['datetime'] = export_df['datetime'].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        # Round RADIUS to 2 decimal places for cleaner display
        export_df['RADIUS'] = export_df['RADIUS'].round(2)
        
        # Reorder columns for better presentation
        column_order = [
            'datetime', 'LotNumber', 'WaferID', 'Slot', 'Condition', 'dmt', 'Label', 'Datum', 
            'XWaferLoc', 'YWaferLoc', 'RADIUS', 'location_id'
        ]
        
        # Only include columns that exist in the dataframe
        available_columns = [col for col in column_order if col in export_df.columns]
        export_df = export_df[available_columns]
        
        # Write to Excel with multiple sheets
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # Export all data
            export_df.to_excel(writer, sheet_name='All Data', index=False)
            
            # Export Layer 1 Thickness data only
            thickness_data = export_df[export_df['Label'] == 'Layer 1 Thickness'].copy()
            if not thickness_data.empty:
                thickness_data.to_excel(writer, sheet_name='Layer 1 Thickness', index=False)
            
            # Export Goodness-of-Fit data only
            gof_data = export_df[export_df['Label'] == 'Goodness-of-Fit'].copy()
            if not gof_data.empty:
                gof_data.to_excel(writer, sheet_name='Goodness-of-Fit', index=False)
            
            # Add metadata sheet
            metadata = pd.DataFrame({
                'Export Information': [
                    'Export Date/Time',
                    'Total Records',
                    'Layer 1 Thickness Records',
                    'Goodness-of-Fit Records',
                    'Total XML Files Processed',
                    'Total Wafers',
                    'Total Lot Numbers',
                    'Total Conditions',
                    'Records with RADIUS data',
                    'Data Description'
                ],
                'Value': [
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    len(export_df),
                    len(export_df[export_df['Label'] == 'Layer 1 Thickness']),
                    len(export_df[export_df['Label'] == 'Goodness-of-Fit']),
                    len(processed_files),
                    len(export_df['WaferID'].unique()) if not export_df.empty else 0,
                    len(export_df['LotNumber'].unique()) if not export_df.empty else 0,
                    len(export_df['Condition'].unique()) if not export_df.empty else 0,
                    len(export_df[export_df['RADIUS'].notna()]),
                    'Complete measurement data with RADIUS calculations, lot numbers, and process conditions'
                ]
            })
            metadata.to_excel(writer, sheet_name='Export Info', index=False)
        
        return filename, True
        
    except Exception as e:
        print(f"Error exporting full data to Excel: {e}")
        return None, False

def get_filtered_dataframe():
    """Return the dataframe (placeholder for filtering functionality)"""
    return df

# Calculate global Layer 1 Thickness range for scaling
thickness_data = df[df['Label'] == 'Layer 1 Thickness']['Datum']
if not thickness_data.empty:
    thickness_min = thickness_data.min()
    thickness_max = thickness_data.max()
    thickness_range = thickness_max - thickness_min
else:
    thickness_min = thickness_max = thickness_range = 0

# Callback for Layer 1 Thickness boxplot
@app.callback(
    Output('thickness-boxplot', 'children'),
    [Input('yscale-dropdown', 'value')]
)
def update_thickness_boxplot(yscale_percent):
    filtered_df = get_filtered_dataframe()
    
    # Calculate thickness range based on filtered data
    thickness_data = filtered_df[filtered_df['Label'] == 'Layer 1 Thickness']['Datum']
    if not thickness_data.empty:
        thickness_min = thickness_data.min()
        thickness_max = thickness_data.max()
        thickness_range = thickness_max - thickness_min
        
        if thickness_range > 0:
            y_min = thickness_min - (yscale_percent * thickness_range)
            y_max = thickness_max + (yscale_percent * thickness_range)
            y_range = [y_min, y_max]
        else:
            y_range = None
    else:
        y_range = None
        
    return make_boxplot('Layer 1 Thickness', y_range, filtered_df)

# Callback for RADIUS vs Thickness plots
@app.callback(
    Output('radius-thickness-plots', 'children'),
    [Input('yscale-dropdown', 'value'),
     Input('trend-legend-radio', 'value')]
)
def update_radius_thickness_plots(yscale_percent, show_trend_legend):
    filtered_df = get_filtered_dataframe()
    
    # Calculate thickness range based on filtered data
    thickness_data = filtered_df[filtered_df['Label'] == 'Layer 1 Thickness']['Datum']
    if not thickness_data.empty:
        thickness_min = thickness_data.min()
        thickness_max = thickness_data.max()
        thickness_range = thickness_max - thickness_min
        
        if thickness_range > 0:
            y_min = thickness_min - (yscale_percent * thickness_range)
            y_max = thickness_max + (yscale_percent * thickness_range)
            y_range = [y_min, y_max]
        else:
            y_range = None
    else:
        y_range = None
        
    return make_radius_thickness_plots(y_range, show_trend_legend, filtered_df)

# Callback for RADIUS vs Thickness by Condition plots
@app.callback(
    Output('radius-thickness-condition-plots', 'children'),
    [Input('yscale-dropdown', 'value'),
     Input('trend-legend-radio', 'value')]
)
def update_radius_thickness_condition_plots(yscale_percent, show_trend_legend):
    filtered_df = get_filtered_dataframe()
    
    # Calculate thickness range based on filtered data
    thickness_data = filtered_df[filtered_df['Label'] == 'Layer 1 Thickness']['Datum']
    if not thickness_data.empty:
        thickness_min = thickness_data.min()
        thickness_max = thickness_data.max()
        thickness_range = thickness_max - thickness_min
        
        if thickness_range > 0:
            y_min = thickness_min - (yscale_percent * thickness_range)
            y_max = thickness_max + (yscale_percent * thickness_range)
            y_range = [y_min, y_max]
        else:
            y_range = None
    else:
        y_range = None
        
    return make_radius_thickness_by_condition_plots(y_range, show_trend_legend, filtered_df)

# Callback for full data Excel export
@app.callback(
    Output('export-full-data-status', 'children'),
    Input('export-full-data-button', 'n_clicks'),
    prevent_initial_call=True
)
def export_full_data_excel(n_clicks):
    if n_clicks:
        filename, success = export_full_data_to_excel()
        if success:
            return html.Div([
                html.P(f"✅ Successfully exported full data to: {filename}", 
                       style={'color': 'green', 'fontWeight': 'bold', 'margin': '10px 0'}),
                html.P("File saved to current directory", 
                       style={'color': 'gray', 'fontSize': '12px'})
            ])
        else:
            return html.Div([
                html.P("❌ Error exporting full data to Excel", 
                       style={'color': 'red', 'fontWeight': 'bold', 'margin': '10px 0'})
            ])
    return html.Div()

# Callback to populate wafer dropdown options
@app.callback(
    Output('contour-wafer-dropdown', 'options'),
    [Input('yscale-dropdown', 'value')]  # Dummy input to trigger callback
)
def update_wafer_dropdown_options(dummy_input):
    filtered_df = get_filtered_dataframe()
    
    # Get unique wafer IDs that have Layer 1 Thickness data with coordinates
    wafer_thickness_data = filtered_df[
        (filtered_df['Label'] == 'Layer 1 Thickness') &
        (filtered_df['XWaferLoc'].notna()) &
        (filtered_df['YWaferLoc'].notna())
    ]
    
    unique_wafers = sorted(wafer_thickness_data['WaferID'].unique())
    
    options = [{'label': wafer_id, 'value': wafer_id} for wafer_id in unique_wafers]
    return options

# Callback to update contour plot based on selected wafer
@app.callback(
    Output('wafer-contour-plot', 'children'),
    [Input('contour-wafer-dropdown', 'value')]
)
def update_wafer_contour_plot(selected_wafer):
    filtered_df = get_filtered_dataframe()
    return make_wafer_contour_plot(selected_wafer, filtered_df)

app.layout = html.Div([
    html.H1("XML Data Analysis Dashboard"),
    
    # Control Panel
    html.Div([
        # YSCALE Control
        html.Div([
            html.Label("YSCALE:", style={'fontWeight': 'bold', 'marginRight': '10px'}),
            dcc.Dropdown(
                id='yscale-dropdown',
                options=[{'label': f'{i}%', 'value': i/100} for i in range(0, 51)],
                value=0.05,  # Default to 5%
                style={'width': '150px', 'display': 'inline-block'}
            )
        ], style={'display': 'inline-block', 'marginRight': '30px'}),
        
        # Trend Legend Control
        html.Div([
            html.Label("Show Trend Line Legend:", style={'fontWeight': 'bold', 'marginRight': '10px'}),
            dcc.RadioItems(
                id='trend-legend-radio',
                options=[
                    {'label': 'Yes', 'value': True},
                    {'label': 'No', 'value': False}
                ],
                value=True,  # Default to True (show legend)
                inline=True,
                style={'display': 'inline-block'}
            )
        ], style={'display': 'inline-block'})
    ], style={'margin': '20px 0', 'padding': '15px', 'backgroundColor': '#f0f0f0', 'borderRadius': '5px'}),
    
    html.H2("Overall Data - Layer 1 Thickness"),
    html.Div(id='thickness-boxplot'),
    
    html.H2("Overall Data - Goodness-of-Fit"),
    make_boxplot('Goodness-of-Fit'),
    
    html.Hr(),
    
    html.H2("Layer 1 Thickness vs Goodness-of-Fit Correlation"),
    make_scatter_plot(),
    
    html.Hr(),
    
    html.H2("Layer 1 Thickness vs RADIUS by WaferID"),
    html.Div(id='radius-thickness-plots'),
    
    html.Hr(),
    
    html.H2("Layer 1 Thickness vs RADIUS by Condition"),
    html.Div(id='radius-thickness-condition-plots'),
    
    html.Hr(),
    
    html.H2("Statistical Summary by WaferID"),
    make_statistical_summary_table(),
    
    html.Hr(),
    
    # Contour Plot Section
    html.H2("Wafer Contour Plot - Layer 1 Thickness"),
    html.Div([
        html.Label("Select Wafer ID:", style={'fontWeight': 'bold', 'marginRight': '10px'}),
        dcc.Dropdown(
            id='contour-wafer-dropdown',
            options=[],  # Will be populated by callback
            value=None,
            placeholder="Select a wafer to view contour plot",
            style={'width': '300px', 'display': 'inline-block'}
        )
    ], style={'margin': '10px 0'}),
    html.Div(id='wafer-contour-plot'),
    
    html.Hr(),
    
    # Export Section
    html.Div([
        html.H3("Export Data to Excel"),
        html.P("Export the complete dataset including all measurements, RADIUS calculations, and process conditions", 
               style={'margin': '0 0 10px 0', 'fontSize': '14px'}),
        html.Button(
            "Export Full Dataset", 
            id="export-full-data-button", 
            n_clicks=0,
            style={
                'backgroundColor': '#28a745',
                'color': 'white',
                'border': 'none',
                'padding': '10px 20px',
                'fontSize': '16px',
                'borderRadius': '5px',
                'cursor': 'pointer',
                'marginBottom': '10px'
            }
        ),
        html.Div(id='export-full-data-status')
    ], style={'margin': '20px 0', 'padding': '15px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'}),
    
    html.Hr(),
    
    html.H2("Processed XML Files"),
    make_files_table()
])

if __name__ == '__main__':
    print(f"DataFrame shape: {df.shape}")
    print(f"Columns: {df.columns.tolist()}")
    print(f"Sample data:")
    if not df.empty:
        print(df.head())
    app.run(debug=True)