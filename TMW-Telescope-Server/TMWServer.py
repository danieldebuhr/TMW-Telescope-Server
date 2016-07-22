# -*- coding: utf-8 -*-
import configparser
import datetime
import io
import os
import socket
import subprocess
import time
import cherrypy
import pythoncom
import struct
import win32com
import win32com.client
from astropy.coordinates import *
from astropy.time import Time
from cherrypy.lib import file_generator
import queue
import threading
import http.client
from cherrypy.process.plugins import SimplePlugin


class BackgroundTaskQueue(SimplePlugin):
    thread = None

    def __init__(self, bus, qsize=100, qwait=2, safe_stop=True):
        SimplePlugin.__init__(self, bus)
        self.q = queue.Queue(qsize)
        self.qwait = qwait
        self.safe_stop = safe_stop

    def start(self):
        self.running = True
        if not self.thread:
            self.thread = threading.Thread(target=self.run)
            self.thread.start()

    def stop(self):
        if self.safe_stop:
            self.running = "draining"
        else:
            self.running = False

        if self.thread:
            self.thread.join()
            self.thread = None
        self.running = False

    def run(self):
        while self.running:
            try:
                try:
                    func, args, kwargs = self.q.get(block=True, timeout=self.qwait)
                except queue.Empty:
                    if self.running == "draining":
                        return
                    continue
                else:
                    func(*args, **kwargs)
                    if hasattr(self.q, 'task_done'):
                        self.q.task_done()
            except:
                self.bus.log("Error in BackgroundTaskQueue %r." % self,
                             level=40, traceback=True)

    def put(self, func, *args, **kwargs):
        """Schedule the given func to be run."""
        self.q.put((func, args, kwargs))

bgtask = BackgroundTaskQueue(cherrypy.engine)
bgtask.subscribe()

class PHDCommunicator():

    def __init__(self):
        self.s = socket.socket()         # Create a socket object
        self.host = socket.gethostname()      # Get local machine name
        self.port = 4300                # Reserve a port for your service.
        self.s.connect((self.host, self.port))

    def __del__(self):
        self.s.close()

    def _sendandreceive(self, cmd):
        byte = bytearray()
        byte.append(cmd)
        self.s.send(byte)
        dat = self.s.recv(8)
        return struct.unpack('B', dat)[0]

    def getstatus(self, noniceout=False):
        status = self._sendandreceive(17)
        if noniceout:
            return status
        else:
            if status == 0:
                return "Leerlauf (not paused, looping, or guiding)"
            elif status == 1:
                return "Loop l&auml;uft und Stern wurde ausgew&auml;hlt"
            elif status == 2:
                return "Kalibrieren..."
            elif status == 3:
                return "Guiding aktiv und Stern ausgew&auml;hlt"
            elif status == 4:
                return "Guiding aktiv aber Stern verloren!"
            elif status == 100:
                return "Pausiert"
            elif status == 101:
                return "Loop l&auml;uft aber kein Stern ausgew&auml;hlt"

    def autoselectstar(self):
        status = self._sendandreceive(14)
        if status == 1:
            return True
        else:
            return False

    def startloop(self):
        self._sendandreceive(19)
        stat = self.getstatus(True)
        if stat == 101 or stat == 1:
            return True
        else:
            return False

    def startguide(self):
        self._sendandreceive(20)
        stat = self.getstatus(True)
        if stat == 2 or stat == 4:
            return True
        else:
            return False

    def stop(self):
        self._sendandreceive(18)
        stat = self.getstatus(True)
        if stat == 0 or stat == 100:
            return True
        else:
            return False


class BYECommunicator():
    def __init__(self):
        try:
            self.s = socket.socket()  # Create a socket object
            self.host = socket.gethostname()  # Get local machine name
            self.port = 1499  # BYE Port
            self.s.connect((self.host, self.port))
        except Exception as e:
            raise(e)

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
        self._send("takepicture quality:raw duration:" + duration + " iso:" + iso + " bin:1")

    def sendconnect(self):
        status = self._sendandreceive("connect")
        return status


class TMWServer(object):
    Config = configparser.ConfigParser()
    Config.read("./config.cfg")

    current_dir = os.path.dirname(os.path.abspath(__file__))

    def responseServer(self, host, port, cmd, key):
        conn = http.client.HTTPConnection(host, port)
        conn.request("GET", "/" + cmd + "/" + key)
        r1 = conn.getresponse()
        conn.close()
        print(r1.status, r1.reason)
        pass

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
        data = f.readall()
        f.close()
        return file_generator(data)

    @cherrypy.expose
    def run(self, name):
        self.bdsrun(name)

    def bdsrun(self, name):
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

    def background_eqmod_goto_name(self, host, port, cmd, key, objekt):
        try:
            print("GoTo: " + objekt)

            now = datetime.datetime.now()
            observatory_location = EarthLocation(lat=53.082806, lon=7.800694, height=5)
            observing_time = Time(now)  # 1am UTC=6pm AZ mountain time
            observer = AltAz(location=observatory_location, obstime=observing_time)

            skyobject = SkyCoord.from_name(objekt)
            skyobject2 = skyobject.transform_to(observer)
            newAltAzcoordiantes = SkyCoord(alt=skyobject2.alt, az=skyobject2.az, obstime=observing_time, frame='altaz',
                                           location=observatory_location)

            pythoncom.CoInitialize()
            o = win32com.client.Dispatch("EQMOD.Telescope")
            o.Unpark()
            o.SlewToCoordinates(str(newAltAzcoordiantes.icrs.ra.hour), str(newAltAzcoordiantes.icrs.dec.degree))

            self.responseServer(host, port, cmd, key)
            return True
        except Exception as e:
            return False


    @cherrypy.expose
    @cherrypy.tools.json_out()
    def eqmod_goto_name(self, host, port, cmd, key, object_name):
        try:
            bgtask.put(self.background_eqmod_goto_name, host, port, cmd, key, object_name)
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': "Fehler beim GoTo", 'detail': str(e)}

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
    @cherrypy.tools.json_out()
    def bye_takepicture(self, duration, iso):
        try:
            status = None
            bye = BYECommunicator()
            bye.takepicture(duration, iso)
            bye = None
            return {'status': True}
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

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def phd_status(self):
        try:
            phd = PHDCommunicator()
            status = phd.getstatus()
            phd = None
            return {'status': True, 'message': status}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def phd_start(self):
        try:
            self.bdsrun("phd_starten")
            time.sleep(30)
            phd = PHDCommunicator()
            status = phd.getstatus()
            phd = None
            return {'status': True, 'message': status}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def phd_guiding_start(self):
        try:
            phd = PHDCommunicator()
            if phd.startloop():
                if phd.autoselectstar():
                    if phd.startguide():
                        return {'status': True, 'message': phd.getstatus()}
                    else:
                        return {'status': False, 'message': phd.getstatus()}
                else:
                    return {'status': False, 'message': "Keinen Stern gefunden, noch mal probieren oder Position korrigieren. PHD: " + phd.getstatus()}
            else:
                return {'status': False,
                        'message': "Fehler: StartLoop klappt nicht. PHD: " + phd.getstatus()}
        except Exception as e:
            phd = None
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def phd_guiding_stop(self):
        try:
            phd = PHDCommunicator()
            if phd.stop():
                phd = None
                return {'status': True, 'message': phd.getstatus()}
            else:
                phd = None
                return {'status': False, 'message': phd.getstatus()}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def at_start(self):
        try:
            self.bdsrun("at_start")
            # todo: Start validate
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def at_platesolve(self):
        try:
            self.bdsrun("at_platesolve")
            # todo: Funktion zur Auswertung von Plate Solve Ergebnissen?
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def bye_start(self):
        try:
            self.bdsrun("bye_start")
            time.sleep(10)
            try:
                status = None
                bye = BYECommunicator()
                status = bye.getstatus()
                bye = None
                return {'status': status != "error"}
            except Exception as e:
                return {'status': False, 'message': str(e)}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def bye_beenden(self):
        try:
            self.bdsrun("bye_beenden")
            # todo: Validieren
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def phd_beenden(self):
        try:
            self.bdsrun("phd_beenden")
            # todo: Validieren
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def shutdown(self):
        try:
            subprocess.call(["shutdown", "-f", "-s", "-t", "10"])
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}


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
