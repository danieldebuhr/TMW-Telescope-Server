# -*- coding: utf-8 -*-
import configparser
import datetime
import io
import os
import socket
import subprocess

import cherrypy
import pythoncom
import win32com
import win32com.client
from astropy.coordinates import *
from astropy.time import Time
from cherrypy.lib import file_generator


class BYECommunicator():
    def __init__(self):
        self.s = socket.socket()  # Create a socket object
        self.host = socket.gethostname()  # Get local machine name
        self.port = 1499  # BYE Port
        self.s.connect((self.host, self.port))

    def __del__(self):
        self.s.close()

    def _sendandreceive(self, cmd):
        self.s.send(cmd.encode())
        dat = self.s.recv(16).decode()
        return dat

    def _send(self, cmd):
        self.s.send(cmd.encode())

    def getstatus(self, noniceout=False):
        status = self._sendandreceive("getstatus")
        return status

    def getpicturepath(self, noniceout=False):
        picturepath = self._sendandreceive("getpicturepath")
        return picturepath

    def takepicture(self, duration, iso):
        self._send("takepicture duration:" + duration + " iso:" + iso + " bin:1")

    def sendconnect(self):
        status = self._sendandreceive("connect")
        return status


class TMWServer(object):
    Config = configparser.ConfigParser()
    Config.read("./config.cfg")

    current_dir = os.path.dirname(os.path.abspath(__file__))

    @cherrypy.expose
    def index(self):

        return """<html>
            <head>
                <script src='static/jquery-3.0.0.min.js'></script>
                <style>
                    #screenshot {
                        max-width: 70%;
                        border: 1px solid gray;
                    }
                </style>
            </head>
            <body>
            <h1>TMWSerer ready</h1>
            <p>""" + Config.get("Settings", "ServerName") + """</p>
            <p>""" + str(datetime.datetime.now()) + """</p>
            <p>
                <img id='screenshot' src='/screenshot' />
            </p>
            <script>
                $("document").ready(function() {
                    setInterval(function() {
                        d = new Date();
                        $("#screenshot").attr('src', '/screenshot?'+d.getTime());
                    }, 10000);
                });
            </script>
            </body>
          </html>"""

    @cherrypy.expose
    def screenshot(self, **params):
        cherrypy.response.headers['Content-Type'] = 'image/png'
        proc = subprocess.Popen(current_dir + "\\" + Config.get("Settings", "PathToScreenshot"))
        proc.wait()
        f = io.open('screenshot.png', 'rb')
        f.seek(0)
        return file_generator(f)

    @cherrypy.expose
    def run(self, name):
        bdsrun = self.current_dir + "\\BDSRun.exe /Script:"
        bdsrun_folder = "\\baramundi\\"
        os.popen(bdsrun + self.current_dir + bdsrun_folder + name + ".bds /S")

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def eqmod_start(self):
        try:
            pythoncom.CoInitialize()
            o = win32com.client.Dispatch("EQMOD.Telescope")
            o.Connected = True
            o.IncClientCount()
            if o.CanSlew:
                return {'status': True}
            else:
                return {'status': False,
                        'message': "Fehler beim Starten von EQMOD (Teleskop verbunden & eingeschaltet?)"}
        except Exception as e:
            return {'status': False, 'message': "Fehler beim Starten von EQMOD: " + e.message}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def eqmod_stop(self):
        try:
            pythoncom.CoInitialize()
            o = win32com.client.Dispatch("EQMOD.Telescope")
            o.StopClientCount()
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': "Konnte EQMOD nicht beenden", 'detail': e.message}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def eqmod_unpark(self):
        try:
            pythoncom.CoInitialize()
            o = win32com.client.Dispatch("EQMOD.Telescope")
            o.Unpark()
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': "Konnte EQMOD nicht unparken", 'detail': e.message}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def eqmod_park(self):
        try:
            pythoncom.CoInitialize()
            o = win32com.client.Dispatch("EQMOD.Telescope")
            o.Park()
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': "Konnte EQMOD nicht parken", 'detail': e.message}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def eqmod_setparkposition(self):
        try:
            pythoncom.CoInitialize()
            o = win32com.client.Dispatch("EQMOD.Telescope")
            o.SetPark()
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': "Konnte EQMOD ParkPosition nicht setzen", 'detail': e.message}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def eqmod_goto_name(self, object_name):
        try:
            now = datetime.datetime.now()
            observatory_location = EarthLocation(lat=53.082806, lon=7.800694, height=5)
            observing_time = Time(now)  # 1am UTC=6pm AZ mountain time
            observer = AltAz(location=observatory_location, obstime=observing_time)

            skyobject = SkyCoord.from_name(object_name)
            skyobject2 = skyobject.transform_to(observer)
            newAltAzcoordiantes = SkyCoord(alt=skyobject2.alt, az=skyobject2.az, obstime=observing_time, frame='altaz',
                                           location=observatory_location)

            pythoncom.CoInitialize()
            o = win32com.client.Dispatch("EQMOD.Telescope")
            o.Unpark()
            o.SlewToCoordinates(str(newAltAzcoordiantes.icrs.ra.hour), str(newAltAzcoordiantes.icrs.dec.degree))

            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': "Konnte EQMOD ParkPosition nicht setzen", 'detail': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def bye_takepicture(self, sec, iso):
        try:
            bye = BYECommunicator()
            bye.takepicture(sec, iso)
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def bye_status(self):
        try:
            status = None
            bye = BYECommunicator()
            status = bye.getstatus()
            bye = None
            return {'status': status != "error", 'message': status}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    def bye_lastpicture(self):
        try:
            cherrypy.response.headers['Content-Type'] = 'application/octet-stream'
            cherrypy.response.headers['Content-Disposition'] = 'attachment; filename="lastpicture.cr2"'
            bye = BYECommunicator()
            image = bye.getpicturepath()
            print(image)
            bye = None
            f = io.open(image, 'rb')
            f.seek(0)
            return file_generator(f)

        except Exception as e:
            cherrypy.response.headers['Content-Type'] = 'application/json'
            return ""

def validate_password(realm, username, password):
    if server_challenge == password:
        return True
    return False


if __name__ == '__main__':
    Config = configparser.ConfigParser()
    Config.read("./config.cfg")
    server_challenge = Config.get("Settings", "ServerChallenge")

    cherrypy.config.update({'tools.auth_basic.checkpassword': validate_password,
                            'tools.auth_basic.on': True,
                            'tools.auth_basic.realm': "localhost"})
    current_dir = os.path.dirname(os.path.abspath(__file__))

    cherrypy.quickstart(TMWServer(), "/", "server.cfg")
