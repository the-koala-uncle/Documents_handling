from http.server import HTTPServer, CGIHTTPRequestHandler  
print('http://192.168.1.8:8080') 
port = 8080
  
httpd = HTTPServer(('', port), CGIHTTPRequestHandler)  
print("Starting simple_httpd on port: " + str(httpd.server_port))  
httpd.serve_forever()  

#cmd  ipconfig
