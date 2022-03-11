from turtle import circle
from server import app as flaskapp
from commonUtils import log
import sys
from renderer import makeWindow
from threading import Thread
import requests
import pygame
import time
import os


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
                flaskapp.run(port=port)
                self.finalport = port
                print('port set')
                return
            
        
        except requests.exceptions.ConnectionError:
            print('ConErr')
            if not serverStatus:
                self.finalport = port
                print('port set')
                flaskapp.run(port=port)

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
        

class LoadingScreen:
    def __init__(self):
        self.exit_ = False
                        
        self.port = 5322
        self.alt_port = 1209

                        
        self.WIDTH = 400
        self.HEIGHT = 300
        self.FPS = 25
        
        pygame.init()

        self.clock = pygame.time.Clock()
        
        self.screen = pygame.display.set_mode((self.WIDTH, self.HEIGHT))
        pygame.display.set_caption("CorelDraw Macro Tools")
        
        self.frames = os.listdir('./loadingscreen')
        
        self.no_frames = len(self.frames)
        
        self.runApp()
        
    def runApp(self):
        self.mainApp = corelMacroGen()
        
        Thread(target=lambda: self.mainApp.run_server(self.port, self.alt_port, n=1), daemon=True).start()



    def bomb(self):
        time.sleep(10)
        self.exit_ = True
    
    
    def loadingScreen(self):
        bomb = Thread(target=self.bomb , daemon=True).start()

        curr_frame = 0
        while not self.exit_:
            self.clock.tick(self.FPS)
            if curr_frame == self.no_frames:
                curr_frame = 0
                
                
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    pygame.quit()
                    
            self.screen.blit(pygame.image.load(f'./loadingscreen/{self.frames[curr_frame]}'), (0, 0))
            
            curr_frame+=1
            pygame.display.update()       




try:
    ls = LoadingScreen()
    ls.loadingScreen()
except:
    pass

pygame.quit()
ls.mainApp.run_renderer()