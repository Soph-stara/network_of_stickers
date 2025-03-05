import pandas as pd
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np

def extract_coordinates(coord_str):
    """Extract latitude and longitude from coordinate string"""
    if isinstance(coord_str, str) and coord_str != 'N/A' and ',' in coord_str:
        try:
            parts = coord_str.split(',')
            if len(parts) >= 2:
                lat = float(parts[0].strip())
                lon = float(parts[1].strip())
                return lat, lon
        except (ValueError, TypeError):
            pass
    return None, None

def create_simple_map(sticker_location_file, general_file):
    """Create a map showing only the sticker locations (Points A-I)"""
    # Load data from Excel files
    sticker_df = pd.read_excel(sticker_location_file)
    
    # Extract unique sticker locations
    unique_points = {}
    for _, row in sticker_df.drop_duplicates(subset=['sticker_location']).iterrows():
        location = row['sticker_location']
        coordinates = row['sticker_location_coordinates']
        if pd.notna(location) and pd.notna(coordinates):
            lat, lon = extract_coordinates(coordinates)
            if lat is not None and lon is not None:
                unique_points[location] = (lat, lon)
    
    # Create a graph
    G = nx.Graph()
    
    # Add nodes (points) to the graph
    for point, (lat, lon) in unique_points.items():
        G.add_node(point, pos=(lon, lat))
    
    # Add edges between points (optional, creates a simple connected path)
    points = list(unique_points.keys())
    for i in range(len(points)-1):
        G.add_edge(points[i], points[i+1])
    
    # Get positions for plotting
    pos = nx.get_node_attributes(G, 'pos')
    
    # Plot the network
    plt.figure(figsize=(12, 10))
    
    try:
        # Try using contextily for map background
        import contextily as ctx
        
        # Create axis
        ax = plt.gca()
        
        # Draw edges
        nx.draw_networkx_edges(G, pos, edge_color='blue', width=2, alpha=0.7, ax=ax)
        
        # Draw nodes
        nx.draw_networkx_nodes(G, pos, node_size=300, node_color='red', alpha=0.8, ax=ax)
        
        # Draw labels - only for sticker locations
        labels = {point: point for point in G.nodes()}
        nx.draw_networkx_labels(G, pos, labels=labels, font_size=12, font_weight='bold',
                              bbox=dict(facecolor='white', alpha=0.7, edgecolor='none', pad=2), ax=ax)
        
        # Set limits to include all points with some padding
        lons = [x[0] for x in pos.values()]
        lats = [x[1] for x in pos.values()]
        lon_margin = 0.003  # About 300 meters at this latitude
        lat_margin = 0.003
        plt.xlim(min(lons) - lon_margin, max(lons) + lon_margin)
        plt.ylim(min(lats) - lat_margin, max(lats) + lat_margin)
        
        # Add title
        plt.title("Sticker Locations (Points A-I)", fontsize=15, pad=20)
        
        # Add the basemap
        ctx.add_basemap(ax, crs='EPSG:4326', source=ctx.providers.OpenStreetMap.Mapnik)
        
        # Save with higher resolution
        plt.savefig('sticker_locations_map.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        print("Map created successfully!")
        
    except ImportError:
        print("Contextily package not installed. Creating basic plot...")
        # Draw edges
        nx.draw_networkx_edges(G, pos, edge_color='gray', alpha=0.6)
        
        # Draw nodes
        nx.draw_networkx_nodes(G, pos, node_size=200, node_color='red', alpha=0.8)
        
        # Draw labels
        nx.draw_networkx_labels(G, pos, font_size=12, font_weight='bold')
        
        # Add title
        plt.title("Sticker Locations (Points A-I)", fontsize=15)
        
        # Add axis labels
        plt.xlabel("Longitude", fontsize=12)
        plt.ylabel("Latitude", fontsize=12)
        
        # Add grid
        plt.grid(True, linestyle='--', alpha=0.6)
        
        plt.tight_layout()
        plt.savefig('sticker_locations_map.png', dpi=300)
        plt.close()
        
    return unique_points

def create_expanded_map(sticker_location_file, general_file):
    """Create a map showing both sticker locations and organization locations"""
    try:
        import folium
        from folium.plugins import MarkerCluster
        
        # Load data
        sticker_df = pd.read_excel(sticker_location_file)
        general_df = pd.read_excel(general_file)
        
        # Create unique points dictionaries
        sticker_points = {}
        for _, row in sticker_df.drop_duplicates(subset=['sticker_location']).iterrows():
            location = row['sticker_location']
            coordinates = row['sticker_location_coordinates']
            if pd.notna(location) and pd.notna(coordinates):
                lat, lon = extract_coordinates(coordinates)
                if lat is not None and lon is not None:
                    sticker_points[location] = (lat, lon)
        
        # Extract valid organization locations
        org_points = {}
        for _, row in general_df.iterrows():
            coords = row['organization_location_coordinates']
            if pd.notna(coords) and coords != 'N/A':
                lat, lon = extract_coordinates(coords)
                if lat is not None and lon is not None:
                    sticker_id = row['sticker_id']
                    org_name = str(row['organization']) if pd.notna(row['organization']) else "Unknown"
                    org_key = f"{org_name}_{sticker_id}"
                    org_points[org_key] = {
                        'coords': (lat, lon),
                        'org_name': org_name,
                        'sticker_id': sticker_id,
                        'topic': str(row['topic']) if pd.notna(row['topic']) else "Unknown"
                    }
        
        # Find connections between stickers and organizations
        connections = []
        for _, sticker_row in sticker_df.iterrows():
            sticker_id = sticker_row['sticker_ID']
            sticker_loc = sticker_row['sticker_location']
            
            # Find matching organization
            for org_key, org_data in org_points.items():
                if org_data['sticker_id'] == sticker_id:
                    if sticker_loc in sticker_points:
                        connections.append({
                            'sticker_loc': sticker_loc,
                            'sticker_coords': sticker_points[sticker_loc],
                            'org_key': org_key,
                            'org_coords': org_data['coords'],
                            'org_name': org_data['org_name'],
                            'topic': org_data['topic']
                        })
        
        # Calculate map center
        all_lats = [lat for lat, _ in sticker_points.values()]
        all_lats.extend([data['coords'][0] for data in org_points.values()])
        
        all_lons = [lon for _, lon in sticker_points.values()]
        all_lons.extend([data['coords'][1] for data in org_points.values()])
        
        center_lat = sum(all_lats) / len(all_lats)
        center_lon = sum(all_lons) / len(all_lons)
        
        # Create map
        m = folium.Map(location=[center_lat, center_lon], zoom_start=14, tiles='OpenStreetMap')
        
        # Add sticker location markers (red)
        for loc, (lat, lon) in sticker_points.items():
            folium.Marker(
                [lat, lon],
                popup=f"<b>{loc}</b>",
                tooltip=loc,
                icon=folium.Icon(color='red', icon='info-sign')
            ).add_to(m)
        
        # Add organization markers (green)
        for org_key, data in org_points.items():
            lat, lon = data['coords']
            folium.Marker(
                [lat, lon],
                popup=f"<b>Organization:</b> {data['org_name']}<br><b>Topic:</b> {data['topic']}",
                tooltip=data['org_name'],
                icon=folium.Icon(color='green', icon='home')
            ).add_to(m)
        
        # Add connections
        for conn in connections:
            sticker_lat, sticker_lon = conn['sticker_coords']
            org_lat, org_lon = conn['org_coords']
            
            folium.PolyLine(
                [(sticker_lat, sticker_lon), (org_lat, org_lon)],
                color='blue',
                weight=2,
                opacity=0.5,
                dash_array='5'
            ).add_to(m)
        
        # Save map
        m.save('expanded_network_map.html')
        print(f"Interactive map created with {len(sticker_points)} sticker locations, {len(org_points)} organizations, and {len(connections)} connections")
        return True
        
    except ImportError as e:
        print(f"Error creating interactive map: {e}")
        return False

if __name__ == "__main__":
    # File paths
    sticker_location_file = '/Users/sophiehamann/Documents/Angewandte_bewerbung_cds/analysis/sticker_location.xlsx'
    general_file = '/Users/sophiehamann/Documents/Angewandte_bewerbung_cds/analysis/general.xlsx'
    
    # Create simple map with just the sticker locations (Points A-I)
    print("Creating map of sticker locations...")
    sticker_points = create_simple_map(sticker_location_file, general_file)
    print(f"Created map with {len(sticker_points)} sticker locations")
    
    # Create expanded interactive map with sticker locations and organizations
    print("\nCreating expanded interactive map...")
    success = create_expanded_map(sticker_location_file, general_file)
    
    if success:
        print("\nVisualization complete!")
        print("Output files:")
        print("1. sticker_locations_map.png - Basic map showing sticker locations (Points A-I)")
        print("2. expanded_network_map.html - Interactive map showing sticker locations and organizations")
    else:
        print("\nCould not create interactive map. Please install folium:")
        print("pip install folium")
    
    print("\nTo ensure proper map background, make sure you have:")
    print("pip install contextily")