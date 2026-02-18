"""
Parser for DOCX chart files to extract chart data and metadata.
"""

import xml.etree.ElementTree as ET
from ..utils.xml_utils import NSMAP


def parse_chart_xml(zipf, chart_path):
    """
    Parses a chart XML file to extract chart information.
    
    Args:
        zipf: ZipFile object of the DOCX file
        chart_path: Path to the chart XML file (e.g., 'word/charts/chart1.xml')
    
    Returns:
        dict: Chart information including title, type, and data points
    """
    chart_info = {
        'title': None,
        'chart_type': None,
        'data_points': [],
        'has_data': False
    }
    
    try:
        chart_xml = zipf.read(chart_path)
        root = ET.fromstring(chart_xml)
        
        # Chart namespace
        chart_ns = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        
        # Get chart title
        title_elem = root.find('.//{' + chart_ns + '}title')
        if title_elem is not None:
            tx_elem = title_elem.find('.//{' + chart_ns + '}tx')
            if tx_elem is not None:
                # Try direct v element first
                v_elem = tx_elem.find('.//{' + chart_ns + '}v')
                if v_elem is not None and v_elem.text:
                    chart_info['title'] = v_elem.text
                else:
                    # Try strCache -> pt -> v structure
                    str_cache = tx_elem.find('.//{' + chart_ns + '}strCache')
                    if str_cache is not None:
                        pt_elems = str_cache.findall('.//{' + chart_ns + '}pt')
                        if pt_elems:
                            # Get first pt element's v text
                            v_elem = pt_elems[0].find('.//{' + chart_ns + '}v')
                            if v_elem is not None and v_elem.text:
                                chart_info['title'] = v_elem.text
        
        # Get chart type
        plot_area = root.find('.//{' + chart_ns + '}plotArea')
        if plot_area is not None:
            # Check for different chart types
            chart_types = [
                ('barChart', 'Bar Chart'),
                ('lineChart', 'Line Chart'),
                ('pieChart', 'Pie Chart'),
                ('areaChart', 'Area Chart'),
                ('scatterChart', 'Scatter Chart'),
                ('bubbleChart', 'Bubble Chart'),
                ('doughnutChart', 'Doughnut Chart'),
                ('radarChart', 'Radar Chart'),
                ('surfaceChart', 'Surface Chart'),
            ]
            
            for chart_tag, chart_name in chart_types:
                if plot_area.find('.//{' + chart_ns + '}' + chart_tag) is not None:
                    chart_info['chart_type'] = chart_name
                    break
        
        # Try to extract data series
        # This is complex and varies by chart type, so we'll do a basic extraction
        for series in root.findall('.//{' + chart_ns + '}ser'):
            series_name = None
            values = []
            categories = []
            
            # Get series name - try multiple structures
            tx_elem = series.find('.//{' + chart_ns + '}tx')
            if tx_elem is not None:
                # Try direct v element
                v_elem = tx_elem.find('.//{' + chart_ns + '}v')
                if v_elem is not None and v_elem.text:
                    series_name = v_elem.text
                else:
                    # Try strCache -> pt -> v structure
                    str_cache = tx_elem.find('.//{' + chart_ns + '}strCache')
                    if str_cache is not None:
                        pt_elems = str_cache.findall('.//{' + chart_ns + '}pt')
                        if pt_elems:
                            v_elem = pt_elems[0].find('.//{' + chart_ns + '}v')
                            if v_elem is not None and v_elem.text:
                                series_name = v_elem.text
            
            # Get categories (x-axis labels)
            cat_elem = series.find('.//{' + chart_ns + '}cat')
            if cat_elem is not None:
                str_cache = cat_elem.find('.//{' + chart_ns + '}strCache')
                if str_cache is not None:
                    for pt in str_cache.findall('.//{' + chart_ns + '}pt'):
                        v_elem = pt.find('.//{' + chart_ns + '}v')
                        if v_elem is not None and v_elem.text:
                            categories.append(v_elem.text)
            
            # Get values (y-axis data)
            val_elem = series.find('.//{' + chart_ns + '}val')
            if val_elem is not None:
                # Try numCache structure (most common)
                num_cache = val_elem.find('.//{' + chart_ns + '}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall('.//{' + chart_ns + '}pt'):
                        idx = pt.get('idx', '0')
                        v_elem = pt.find('.//{' + chart_ns + '}v')
                        if v_elem is not None and v_elem.text:
                            try:
                                value = float(v_elem.text)
                                values.append(value)
                                chart_info['has_data'] = True
                            except ValueError:
                                pass
                
                # Also check numLit (less common)
                num_lit = val_elem.find('.//{' + chart_ns + '}numLit')
                if num_lit is not None and not values:
                    for pt in num_lit.findall('.//{' + chart_ns + '}pt'):
                        idx = pt.get('idx', '0')
                        v_elem = pt.find('.//{' + chart_ns + '}v')
                        if v_elem is not None and v_elem.text:
                            try:
                                value = float(v_elem.text)
                                values.append(value)
                                chart_info['has_data'] = True
                            except ValueError:
                                pass
            
            # Store series data if we found any
            if series_name or values:
                series_data = {
                    'series_name': series_name or 'Unnamed Series',
                    'values': values,
                    'categories': categories if categories else None
                }
                chart_info['data_points'].append(series_data)
        
        return chart_info
    except Exception:
        return chart_info


def parse_all_charts(zipf):
    """
    Parses all chart files in a DOCX document.
    
    Args:
        zipf: ZipFile object of the DOCX file
    
    Returns:
        dict: Mapping of chart relationship ID to chart information
    """
    charts = {}
    
    try:
        # Get chart relationships from document relationships
        rels_xml = zipf.read('word/_rels/document.xml.rels')
        rels_root = ET.fromstring(rels_xml)
        
        rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        chart_rel_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
        
        chart_rels = {}
        for rel in rels_root.findall('.//{' + rel_ns + '}Relationship'):
            rel_id = rel.get('Id')
            rel_type = rel.get('Type', '')
            target = rel.get('Target', '')
            
            if chart_rel_type in rel_type or 'chart' in rel_type.lower():
                # Convert relative path to absolute path in ZIP
                chart_path = target if target.startswith('word/') else f'word/{target}'
                chart_rels[rel_id] = chart_path
        
        # Parse each chart
        for rel_id, chart_path in chart_rels.items():
            chart_info = parse_chart_xml(zipf, chart_path)
            charts[rel_id] = chart_info
        
    except Exception:
        pass
    
    return charts
