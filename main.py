from server import app
from commonUtils import log
import sys
from renderer import makeWindow
from threading import Thread
import requests
import time

finalPort = 1234

class corelMacroGen:
    
    finalport = None

    def run_server(self, port, alt_port, n):

        
        if n>4:
            log('Fatal: Port in use')
            sys.exit()
        
        serverStatus = None

        try:
            isServerRunning = requests.get(f'http://localhost:{port}/testing').text
            serverStatus = requests.get(f'http://localhost:{port}/testing').status_code
            print(serverStatus)
            print(isServerRunning)
        except ConnectionRefusedError:
            print('ConnRefErr')
            if not serverStatus:
                app.run(port=port)
                self.finalport = port
                print('port set')
                return
            
        
        except requests.exceptions.ConnectionError:
            print('ConErr')
            if not serverStatus:
                self.finalport = port
                print('port set')
                app.run(port=port)

                return
            

        
        except Exception as e:
            log(str(e))

            return
        

        try:
        
            if isServerRunning == 'asjihdiohdua huei9wa eajsnjaiehw8 aihdajhw0d a90oaidji2ojeiojd aouw9-e a90uw09 assdhh':
                log('Server is running')
                self.finalport = port
                print('port set')
                return
            
            if isServerRunning != 'asjihdiohdua huei9wa eajsnjaiehw8 aihdajhw0d a90oaidji2ojeiojd aouw9-e a90uw09 assdhh' or serverStatus == 404:
                log('Another application is running on the required port.')
                self.finalport = port
                print('port set')
                self.run_server(alt_port, port, n+1)
                
        except Exception as e:
            log(str(e))

            return
                
                
            
    def run_renderer(self):
        print(self.finalport)
        try:
            makeWindow(self.finalport)
            
        except Exception as e:
            print('e')
            log(str(e))
        


port = 5322
alt_port = 1209

mainApp = corelMacroGen()

server = Thread(target=lambda: mainApp.run_server(port, alt_port, n=1), daemon=True).start()

time.sleep(1)

mainApp.run_renderer()

