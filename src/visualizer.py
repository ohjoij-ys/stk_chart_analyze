import plotly.graph_objects as go
from config import COLORS, CHARTS

def format_currency(value):
    """Format value to 亿元 string."""
    if value is None:
        return ""
    # Convert to 100 million
    val_in_yi = value / 100000000
    return f"{val_in_yi:.2f}亿"

def create_chart(df, chart_config, transparent_bg=True):
    """
    Create a plotly figure based on config.
    """
    fig = go.Figure()
    
    # Layout configuration
    layout_args = {
        'template': 'plotly_dark',
        'paper_bgcolor': 'rgba(0,0,0,0)' if transparent_bg else None,
        'plot_bgcolor': 'rgba(0,0,0,0)' if transparent_bg else None,
        'font': {'color': COLORS['grey'], 'family': 'sans-serif'},
        'title': {
            'text': f"<b>{chart_config['name']}</b>", 
            'x': 0.0,
            'font': {'size': 24, 'color': COLORS['white']}
        },
        'xaxis': {
            'showgrid': True, 
            'gridcolor': 'rgba(255, 255, 255, 0.1)', # Subtle grid
            'title': '报告日期',
            'showline': True,
            'linecolor': 'rgba(255, 255, 255, 0.2)'
        },
        'yaxis': {
            'showgrid': True, 
            'gridcolor': 'rgba(255, 255, 255, 0.1)',
            'title': chart_config.get('y_axis_label', ''),
            'zeroline': False
        },
        'legend': {
            'orientation': 'h',
            'yanchor': 'bottom',
            'y': 1.05,
            'xanchor': 'right',
            'x': 1,
            'bgcolor': 'rgba(0,0,0,0)'
        },
        'hovermode': 'x unified',
        'margin': dict(l=20, r=20, t=80, b=20)
    }

    # Add traces
    for series in chart_config['series']:
        field_name = series['field']
        
        if field_name not in df.columns:
            continue
            
        # Data preparation
        # Filter out NaNs for the line to be continuous or handle gaps?
        # Plotly handles NaNs by breaking the line, which is usually correct for missing data.
        
        is_percent = series.get('is_percent', False)
        
        if is_percent:
            y_values = df[field_name]
            hover_template = '%{y:.2f}%<extra></extra>'
        elif field_name in ['pe_ttm', 'pe_mean', 'pe_q30', 'pe_q70', 'ps_ttm', 'ps_mean', 'ps_q30', 'ps_q70']:
            # No division by 1e8 for PE/PS ratio
            y_values = df[field_name]
            hover_template = '%{y:.2f}<extra></extra>'
        else:
            # Scale to 亿元
            y_values = df[field_name] / 100000000
            hover_template = '%{y:.2f}亿<extra></extra>'
        
        trace_args = {
            'x': df.index,
            'y': y_values,
            'name': series['name'],
            'hovertemplate': hover_template
        }
        
        # Determine chart type
        chart_type = series.get('type', 'line') # default to line
        
        if chart_type == 'line':
            line_dict = dict(color=series['color'], width=2)
            if 'dash' in series:
                line_dict['dash'] = series['dash']
                
            trace_args.update({
                'mode': 'lines+markers' if series.get('show_values', False) else 'lines',
                'line': line_dict,
                # 'marker': dict(size=6) # removed default marker
            })
            
            # Support showing values for lines
            if series.get('show_values', False):
                if is_percent:
                    text_values = [f"{v:.2f}%" for v in y_values]
                elif field_name in ['pe_ttm', 'ps_ttm']: # Special handling for PE/PS ratio which is just a number
                    text_values = [f"{v:.2f}" for v in y_values]
                else:
                    text_values = [f"{v:.2f}" for v in y_values]
                
                trace_args.update({
                    'text': text_values,
                    'textposition': 'top center',
                    'textfont': {'color': COLORS['white']},
                    'mode': 'lines+markers+text'
                })

            trace = go.Scatter(**trace_args)
            
        elif chart_type == 'bar':
            # Support per-bar coloring if 'color_field' is present
            if 'color_field' in series and series['color_field'] in df.columns:
                trace_args.update({
                    'marker_color': df[series['color_field']]
                })
            else:
                trace_args.update({
                    'marker_color': series.get('color', COLORS['blue'])
                })
            
            # Support showing values on top of bars
            if series.get('show_values', False):
                if is_percent:
                    text_values = [f"{v:.2f}%" for v in y_values]
                else:
                    text_values = [f"{v:.2f}" for v in y_values]
                
                trace_args.update({
                    'text': text_values,
                    'textposition': 'outside',
                    'textfont': {'color': COLORS['white']}
                })

            trace = go.Bar(**trace_args)
            
        # Handle Dual Axis
        if 'y_axis' in series and series['y_axis'] == 'y2':
            trace.yaxis = 'y2'
            
        fig.add_trace(trace)

    # Add second Y-axis if needed
    if chart_config.get('dual_axis'):
        layout_args['yaxis2'] = {
            'title': chart_config.get('y_axis_2_label', ''),
            'overlaying': 'y',
            'side': 'right',
            'showgrid': False, # Hide grid for second axis to avoid clutter
            'zeroline': False,
            'rangemode': 'tozero',
            'title_font': {'color': COLORS['grey']},
            'tickfont': {'color': COLORS['grey']}
        }

    fig.update_layout(**layout_args)
    
    return fig
