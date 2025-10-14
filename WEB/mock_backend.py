"""
Desktop Management Mock Backend Server
Simulates gdpmappercb.nomura.com IIS SOAP web services
Uses CSV files for data storage (no database required)

Services:
- ClassicMapper.asmx (GetUserDrives, GetUserPrinters, GetUserPersonalFolders)
- ClassicInventory.asmx (Insert* methods)

Usage:
    python mock_backend.py

Server runs on: http://0.0.0.0:80
Configure DNS: gdpmappercb.nomura.com -> <this-server-ip>
"""

from flask import Flask, request, Response
import csv
import os
from datetime import datetime
from xml.etree import ElementTree as ET

app = Flask(__name__)

# Data directory
DATA_DIR = os.path.join(os.path.dirname(__file__), 'Data')
os.makedirs(DATA_DIR, exist_ok=True)

# ============================================================================
# Helper Functions
# ============================================================================

def create_soap_response(method_name, result_xml):
    """Create SOAP envelope response"""
    soap_response = f'''<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" 
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
        <{method_name}Response xmlns="http://webtools.japan.nom">
            <{method_name}Result>{result_xml}</{method_name}Result>
        </{method_name}Response>
    </soap:Body>
</soap:Envelope>'''
    return soap_response

def parse_soap_request(xml_data):
    """Parse incoming SOAP request and extract parameters"""
    try:
        root = ET.fromstring(xml_data)
        # Remove namespaces for easier parsing
        for elem in root.iter():
            if '}' in elem.tag:
                elem.tag = elem.tag.split('}', 1)[1]
        
        # Find Body element
        body = root.find('.//Body')
        if body is None:
            return None, {}
        
        # Get method name and parameters
        method = list(body)[0]
        method_name = method.tag
        
        params = {}
        for child in method:
            params[child.tag] = child.text or ''
        
        return method_name, params
    except Exception as e:
        print(f"Error parsing SOAP: {e}")
        return None, {}

def load_csv(filename):
    """Load CSV file as list of dictionaries"""
    filepath = os.path.join(DATA_DIR, filename)
    if not os.path.exists(filepath):
        return []
    
    with open(filepath, 'r', newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        return list(reader)

def append_csv(filename, data):
    """Append data to CSV file"""
    filepath = os.path.join(DATA_DIR, filename)
    file_exists = os.path.exists(filepath)
    
    with open(filepath, 'a', newline='', encoding='utf-8') as f:
        if data:
            writer = csv.DictWriter(f, fieldnames=data.keys())
            if not file_exists:
                writer.writeheader()
            writer.writerow(data)

def log_request(service, method, params):
    """Log incoming request"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{timestamp}] {service} -> {method}")
    for key, value in params.items():
        print(f"  {key}: {value}")

# ============================================================================
# Mapper Service Endpoints (ClassicMapper.asmx)
# ============================================================================

@app.route('/ClassicMapper.asmx', methods=['POST'])
def mapper_service():
    """Handle ClassicMapper.asmx SOAP requests"""
    xml_data = request.data.decode('utf-8')
    method_name, params = parse_soap_request(xml_data)
    
    if not method_name:
        return Response("Invalid SOAP request", status=400)
    
    log_request('ClassicMapper', method_name, params)
    
    if method_name == 'TestService':
        result = '<TestServiceResult>OK</TestServiceResult>'
        response = create_soap_response('TestService', result)
        return Response(response, mimetype='text/xml')
    
    elif method_name == 'GetUserDrives':
        drives = load_csv('drives.csv')
        user_id = params.get('UserId', '')
        domain = params.get('Domain', '')
        
        # Filter drives for this user (simple matching)
        user_drives = [d for d in drives if d.get('UserId', '').upper() == user_id.upper()]
        
        # Build XML result
        drive_xml = ''
        for drive in user_drives:
            drive_xml += f'''
            <Drive>
                <Id>{drive.get('Id', '1')}</Id>
                <Domain>{drive.get('Domain', domain)}</Domain>
                <UserId>{drive.get('UserId', user_id)}</UserId>
                <AdGroup>{drive.get('AdGroup', '')}</AdGroup>
                <Site>{drive.get('Site', '')}</Site>
                <Drive>{drive.get('Drive', 'H:')}</Drive>
                <UncPath>{drive.get('UncPath', '')}</UncPath>
                <Description>{drive.get('Description', '')}</Description>
                <DisconnectOnLogin>{drive.get('DisconnectOnLogin', 'false')}</DisconnectOnLogin>
            </Drive>'''
        
        result = f'<Drives>{drive_xml}</Drives>' if drive_xml else '<Drives />'
        response = create_soap_response('GetUserDrives', result)
        return Response(response, mimetype='text/xml')
    
    elif method_name == 'GetUserPrinters':
        printers = load_csv('printers.csv')
        computer_name = params.get('HostName', '')
        
        # Filter printers for this computer
        computer_printers = [p for p in printers if p.get('HostName', '').upper() == computer_name.upper()]
        
        # Build XML result
        printer_xml = ''
        for printer in computer_printers:
            printer_xml += f'''
            <Printer>
                <Id>{printer.get('Id', '1')}</Id>
                <UncPath>{printer.get('UncPath', '')}</UncPath>
                <IsDefault>{printer.get('IsDefault', 'false')}</IsDefault>
                <Description>{printer.get('Description', '')}</Description>
            </Printer>'''
        
        result = f'<Printers>{printer_xml}</Printers>' if printer_xml else '<Printers />'
        response = create_soap_response('GetUserPrinters', result)
        return Response(response, mimetype='text/xml')
    
    elif method_name == 'GetUserPersonalFolders':
        psts = load_csv('pst.csv')
        user_id = params.get('UserId', '')
        
        # Filter PSTs for this user
        user_psts = [p for p in psts if p.get('UserId', '').upper() == user_id.upper()]
        
        # Build XML result
        pst_xml = ''
        for pst in user_psts:
            pst_xml += f'''
            <PersonalFolder>
                <Id>{pst.get('Id', '1')}</Id>
                <UserId>{pst.get('UserId', user_id)}</UserId>
                <UncPath>{pst.get('UncPath', '')}</UncPath>
                <DisconnectOnLogin>{pst.get('DisconnectOnLogin', 'false')}</DisconnectOnLogin>
            </PersonalFolder>'''
        
        result = f'<PersonalFolders>{pst_xml}</PersonalFolders>' if pst_xml else '<PersonalFolders />'
        response = create_soap_response('GetUserPersonalFolders', result)
        return Response(response, mimetype='text/xml')
    
    else:
        return Response(f"Unknown method: {method_name}", status=400)

# ============================================================================
# Inventory Service Endpoints (ClassicInventory.asmx)
# ============================================================================

@app.route('/ClassicInventory.asmx', methods=['POST'])
def inventory_service():
    """Handle ClassicInventory.asmx SOAP requests"""
    xml_data = request.data.decode('utf-8')
    method_name, params = parse_soap_request(xml_data)
    
    if not method_name:
        return Response("Invalid SOAP request", status=400)
    
    log_request('ClassicInventory', method_name, params)
    
    if method_name == 'TestService':
        result = '<TestServiceResult>OK</TestServiceResult>'
        response = create_soap_response('TestService', result)
        return Response(response, mimetype='text/xml')
    
    elif method_name == 'InsertLogonInventory':
        data = {
            'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'EventType': 'Logon',
            'UserId': params.get('UserId', ''),
            'UserDomain': params.get('UserDomain', ''),
            'HostName': params.get('HostName', ''),
            'Domain': params.get('Domain', ''),
            'SiteName': params.get('SiteName', ''),
            'City': params.get('City', ''),
            'OuMapping': params.get('OuMapping', '')
        }
        append_csv('sessions.csv', data)
        
        result = '<InsertLogonInventoryResult>SUCCESS</InsertLogonInventoryResult>'
        response = create_soap_response('InsertLogonInventory', result)
        return Response(response, mimetype='text/xml')
    
    elif method_name == 'InsertLogoffInventory':
        data = {
            'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'EventType': 'Logoff',
            'UserId': params.get('UserId', ''),
            'UserDomain': params.get('UserDomain', ''),
            'HostName': params.get('HostName', ''),
            'Domain': params.get('Domain', ''),
            'SiteName': params.get('SiteName', ''),
            'City': params.get('City', ''),
            'OuMapping': params.get('OuMapping', '')
        }
        append_csv('sessions.csv', data)
        
        result = '<InsertLogoffInventoryResult>SUCCESS</InsertLogoffInventoryResult>'
        response = create_soap_response('InsertLogoffInventory', result)
        return Response(response, mimetype='text/xml')
    
    elif method_name == 'InsertActiveDriveMappingsFromInventory':
        data = {
            'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'UserId': params.get('UserId', ''),
            'HostName': params.get('HostName', ''),
            'Domain': params.get('Domain', ''),
            'SiteName': params.get('SiteName', ''),
            'City': params.get('City', ''),
            'Drive': params.get('Drive', ''),
            'UncPath': params.get('UncPath', ''),
            'Description': params.get('Description', ''),
            'OuMapping': params.get('OuMapping', '')
        }
        append_csv('inventory_drives.csv', data)
        
        result = '<InsertActiveDriveMappingsFromInventoryResult>SUCCESS</InsertActiveDriveMappingsFromInventoryResult>'
        response = create_soap_response('InsertActiveDriveMappingsFromInventory', result)
        return Response(response, mimetype='text/xml')
    
    elif method_name == 'InsertMapperPrinterInventory':
        data = {
            'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'UserId': params.get('UserId', ''),
            'HostName': params.get('HostName', ''),
            'Domain': params.get('Domain', ''),
            'UncPath': params.get('UncPath', ''),
            'IsDefault': params.get('IsDefault', ''),
            'Driver': params.get('Driver', ''),
            'Port': params.get('Port', ''),
            'Description': params.get('Description', '')
        }
        append_csv('inventory_printers.csv', data)
        
        result = '<InsertMapperPrinterInventoryResult>SUCCESS</InsertMapperPrinterInventoryResult>'
        response = create_soap_response('InsertMapperPrinterInventory', result)
        return Response(response, mimetype='text/xml')
    
    elif method_name == 'InsertActivePersonalFolderMappingsFromInventory':
        data = {
            'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'UserId': params.get('UserId', ''),
            'HostName': params.get('HostName', ''),
            'Path': params.get('Path', ''),
            'UncPath': params.get('UncPath', ''),
            'Size': params.get('Size', ''),
            'PstLastUpdate': params.get('PstLastUpdate', '')
        }
        append_csv('inventory_pst.csv', data)
        
        result = '<InsertActivePersonalFolderMappingsFromInventoryResult>SUCCESS</InsertActivePersonalFolderMappingsFromInventoryResult>'
        response = create_soap_response('InsertActivePersonalFolderMappingsFromInventory', result)
        return Response(response, mimetype='text/xml')
    
    else:
        return Response(f"Unknown method: {method_name}", status=400)

# ============================================================================
# Health Check
# ============================================================================

@app.route('/')
def index():
    """Simple health check page"""
    return '''
    <html>
    <head><title>Mock Desktop Management Backend</title></head>
    <body>
        <h1>Desktop Management Mock Backend Server</h1>
        <p>Status: <strong style="color: green;">Running</strong></p>
        <h2>Available Services:</h2>
        <ul>
            <li><a href="/ClassicMapper.asmx">/ClassicMapper.asmx</a> (Mapper Service)</li>
            <li><a href="/ClassicInventory.asmx">/ClassicInventory.asmx</a> (Inventory Service)</li>
        </ul>
        <h2>Data Storage:</h2>
        <p>CSV files in: <code>./Data/</code></p>
        <h2>DNS Configuration:</h2>
        <p>Point <code>gdpmappercb.nomura.com</code> to this server's IP address</p>
    </body>
    </html>
    '''

# ============================================================================
# Main
# ============================================================================

if __name__ == '__main__':
    print("=" * 70)
    print("Desktop Management Mock Backend Server")
    print("=" * 70)
    print(f"Data directory: {DATA_DIR}")
    print("Services:")
    print("  - ClassicMapper.asmx (Mapper)")
    print("  - ClassicInventory.asmx (Inventory)")
    print("")
    print("Configure DNS to point gdpmappercb.nomura.com to this server")
    print("=" * 70)
    print("")
    
    # Run server on port 80 (HTTP)
    # Note: On Windows, may need admin rights to bind to port 80
    # Alternative: Use port 8080 and configure hosts file
    app.run(host='0.0.0.0', port=80, debug=True)

