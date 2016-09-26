# -*- coding: utf-8 -*-
import configparser
import datetime
from io import BytesIO
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
    """
    Die Klasse BackgroundTaskQueue ermöglicht die Funktionalität von Rückantworten nach Abschluss einer Aktion.
    http://tools.cherrypy.org/wiki/BackgroundTaskQueue
    """
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


class PHDCommunicator():
    """
    Kommunikations-Klasse für PHD Vebrindungen.
    """

    def __init__(self):
        """
        Konstruktor, baut eine Verbindung auf.
        """

        self.s = socket.socket()
        self.host = socket.gethostname()
        self.port = 4300  # PHD Port
        self.s.connect((self.host, self.port))

    def __del__(self):
        """
        Schließt die Verbindung bei Zerstörung des Kommunikators.
        :return:
        """
        self.s.close()

    def _sendandreceive(self, cmd):
        """
        Sendet einen Befehl an den PHD Server und liest dessen Antwort.
        Mehr Information zu den Befehlen:
        https://github.com/OpenPHDGuiding/phd2/wiki/SocketServerInterface
        :param cmd: Befehl in Form einer Zahl.
        :return: Antwort des Servers in Form einer Zahl.
        """
        byte = bytearray()
        byte.append(cmd)
        self.s.send(byte)
        dat = self.s.recv(8)
        return struct.unpack('B', dat)[0]

    def getstatus(self, returnstatuscode=False):
        """
        Wrapper f+r den Befehl "MSG_GETSTATUS".
        :param returnstatuscode: Bei False (default) menschlich lesbare Antwort.
        :return: Status
        """
        status = self._sendandreceive(17)
        if returnstatuscode:
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
        """
        Wrapper für den Befehl "MSG_AUTOFINDSTAR".
        :return: True/False
        """
        status = self._sendandreceive(14)
        if status == 1:
            return True
        else:
            return False

    def startloop(self):
        """
        Wrapper für den Befehl "MSG_LOOP".
        :return: True/False
        """
        self._sendandreceive(19)
        stat = self.getstatus(True)
        if stat == 101 or stat == 1:
            return True
        else:
            return False

    def startguide(self):
        """
        Wrapper für dem Befehl "MSG_STARTGUIDING".
        :return: True/False
        """
        self._sendandreceive(20)
        stat = self.getstatus(True)
        if stat == 2 or stat == 4:
            return True
        else:
            return False

    def stop(self):
        """
        Wrapper für den Befehl "MSG_STOP".
        :return: True/False
        """
        self._sendandreceive(18)
        stat = self.getstatus(True)
        if stat == 0 or stat == 100:
            return True
        else:
            return False


class BYECommunicator():
    """
    Kommunikator für die Software BackyardEOS.
    """

    def __init__(self):
        """
        Konstruktor, verbindet sich mit BYE.
        """
        try:
            self.s = socket.socket()
            self.s.connect(('localhost', 1499))
        except Exception as e:
            raise (e)

    def __del__(self):
        """
        Schließt die Verbindung bei Software-Ende.
        :return:
        """
        self.s.close()

    def _sendandreceive(self, cmd):
        """
        Sendet einen Befehl an BYE und gibt die Antwort zurück.
        :param cmd: Befehl als String
        :return: Daten aus BYE
        """
        self.s.send(cmd.encode())
        dat = self.s.recv(16).decode()
        return dat

    def _send(self, cmd):
        """
        Sendet einen Befehl, wertet die Antwort nicht aus.
        :param cmd:
        :return:
        """
        self.s.send(cmd.encode())

    def getstatus(self):
        """
        Ermittelt den Status von BYE.
        :return: Status
        """
        status = self._sendandreceive("getstatus")
        return status

    def getpicturepath(self):
        """
        Ermittelt den Pfad des letzten Bildes.
        :return: Pfad des letzten Bildes.
        """
        picturepath = self._sendandreceive("getpicturepath")
        return picturepath

    def takepicture(self, duration, iso):
        """
        Wrapper des Befehls 'takepicture', welches eine Aufnahme startet.
        :param duration:
        :param iso:
        :return:
        """
        self._send("takepicture quality:raw duration:" + duration + " iso:" + iso + " bin:1")

    def sendconnect(self):
        """
        Initialisiert eine Verbindung.
        :return: Status
        """
        status = self._sendandreceive("connect")
        return status


class TMWServer(object):
    """
    Die Klasse TMWServer stellt den eigentlichen lokalen Server dar.
    Hier sind alle Funktionen nach außen präsentiert, die für die
    Steuerung der Software benötigt wird.
    """

    # todo: Mehr Ergebnis-Validierung der einzelnen Funktionen.

    # Config lesen
    Config = configparser.ConfigParser()
    Config.read("./config.cfg")

    # <editor-fold desc="Allgemeine Routen">

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def shutdown(self):
        """
        Route '/shutdown' - fährt den Hostcomputer herunter.
        :return: JSON-Daten des Status.
        """
        try:
            subprocess.call(["shutdown", "-f", "-s", "-t", "10"])
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    def index(self):
        """
        Index-Funktion für die Route '/'. Gibt eine Info-Seite mit Screenshot aus.
        :return:
        """

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
        """
        Route '/screenshot' - Liefert ein Screenshot vom ausführenden Server.
        :param params: -
        :return: PNG-Daten des Screenshots.
        """
        # Header auf image/png für die Bildausgabe setzen
        cherrypy.response.headers['Content-Type'] = 'image/png'
        # Über die mitgelieferte Screenshto.exe ein Bildschirmfoto schießen.
        proc = subprocess.Popen(current_dir + "\\" + Config.get("Settings", "PathToScreenshot"))
        proc.wait()
        # Bildschirmfoto öffnen und per file_generator an den Clienten zurückgeben.
        f = open('screenshot.png', 'rb')
        data = BytesIO(f.read())
        f.close()
        return file_generator(data)

    @cherrypy.expose
    def run(self, name):
        """
        Route '/run' - führt ein BDS-Script aus.
        :param name: Name des Scriptes.
        :return:
        """
        bdsrun(name)
        pass

    # </editor-fold>

    # <editor-fold desc="EQMod Routen">

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def eqmod_start(self):
        """
        Route '/eqmod_start' - startet die Software EQMod.
        :return: JSON-Daten mit Status.
        """
        try:
            pythoncom.CoInitialize()
            # COM-Objekt ansteuern
            o = win32com.client.Dispatch("EQMOD.Telescope")
            # Verbindung mit dem Teleskop herstellen
            o.Connected = True
            o.IncClientCount()
            # Wenn CanSlew True zurückgibt, wurde die Verbindung hergestellt.
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
        """
        Route '/eqmod_stop' - beendet die Software EQMod.
        :return: JSON-Daten mit Status.
        """
        try:
            # COM-Objekt ansteuern
            pythoncom.CoInitialize()
            o = win32com.client.Dispatch("EQMOD.Telescope")
            # Beendet EQMod
            o.StopClientCount()
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': "Konnte EQMOD nicht beenden", 'detail': e.message}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def eqmod_unpark(self):
        """
        Route 'eqmod_unpark' - beendet den Parkmodus von EQMod
        :return: JSON-Daten mit Status.
        """
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
        """
        Route 'eqmod_park' - Parkt das Teleskop.
        :return: JSON-Daten mit Status.
        """
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
        """
        Route '/setparkposition' -  erkennt die aktuelle Position und speichert
        diese als neue Parkposition.
        :return: JSON-Daten mit Status.
        """
        try:
            pythoncom.CoInitialize()
            o = win32com.client.Dispatch("EQMOD.Telescope")
            o.SetPark()
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': "Konnte EQMOD ParkPosition nicht setzen", 'detail': e.message}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def eqmod_goto_name(self, object_name, host="", port="", cmd="", key=""):
        """
        Route '/eqmod_goto_name' - GoTo-Funktion mit Parametern der Rückantwort.
        :param host: Hostname des Servers
        :param port: Port des Servers
        :param cmd: Rückrouten-Befehl
        :param key: Rückrouten-Schlüssel
        :param object_name: GoTo Objekt Name als NGC-Katalogeintrag.
        :return: JSON-Daten mit Status.
        """
        try:
            # Das eigentliche GoTo wird später erledigt. Hier wurde der Befehl nut entgegengenommen
            # und direkt eine Antwort formuliert.
            bgtask.put(background_eqmod_goto_name, object_name, host, port, cmd, key)
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': "Fehler beim GoTo", 'detail': str(e)}

    # </editor-fold>

    # <editor-fold desc="BackyardEOS Routen">

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def bye_start(self):
        """
        Route '/bye_start' - startet BackyardEOS und verbindet sich zum Test.
        :return: JSON-Daten des Status.
        """
        try:
            bdsrun("bye_start")
            time.sleep(20)
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
    def bye_status(self):
        """
        Route 'bye_status' - Ermittelt den aktuellen Status von BYE und gibt diesen aus.
        :return: JSON-Daten mti BYE Status.
        """
        try:
            # Kommunikator erstellen.
            bye = BYECommunicator()
            status = bye.getstatus()
            # Kommunikator wieder beenden.
            bye = None
            return {'status': status != "error", 'message': status}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def bye_takepicture(self, duration, iso):
        """
        Route 'bye_takepicture' - erstellt ein Bild mit den angegebenen Parametern.
        :param duration: Belichtungszeit in Sekunden
        :param iso: ISO-Wert
        :return: JSON-Daten mit Status.
        """
        try:
            bye = BYECommunicator()
            bye.takepicture(duration, iso)
            bye = None
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    def bye_lastpicture(self):
        """
        Route 'bye_lastpicture' - ermittelt das letzte aufgenommene Bild und sendet dieses.
        :return: Bild-Daten
        """
        try:
            cherrypy.response.headers['Content-Type'] = 'application/octet-stream'
            cherrypy.response.headers['Content-Disposition'] = 'attachment; filename="lastpicture.cr2"'
            bye = BYECommunicator()
            image = bye.getpicturepath()
            print(image)
            bye = None

            f = open(image, 'rb')
            data = BytesIO(f.read())
            f.close()

            return file_generator(data)

        except Exception as e:
            cherrypy.response.headers['Content-Type'] = 'application/json'
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def bye_beenden(self):
        """
        Route '/bye_beenden' - beendet BackyardEOS.
        :return: JSON-Daten des Status.
        """
        try:
            bdsrun("bye_beenden")
            # todo: Validieren
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    # </editor-fold>

    # <editor-fold desc="PHD Routen">

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def phd_start(self):
        """
        Route '/phd_start' - startet PHD und testet die Verbindung.
        :return:  JSON-Daten des Status.
        """
        try:
            bdsrun("phd_starten")
            time.sleep(30)
            phd = PHDCommunicator()
            status = phd.getstatus()
            phd = None
            return {'status': True, 'message': status}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def phd_status(self):
        """
        Route '/phd_status' - ermittelt den Status von PHD.
        :return: JSON-Daten des Status.
        """
        try:
            phd = PHDCommunicator()
            status = phd.getstatus()
            phd = None
            return {'status': True, 'message': status}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def phd_guiding_start(self):
        """
        Route '/phd_guiding_start' - startet das Guiding.
        :return: JSON-Daten des Status.
        """
        try:
            phd = PHDCommunicator()
            if phd.startloop():
                if phd.autoselectstar():
                    if phd.startguide():
                        return {'status': True, 'message': phd.getstatus()}
                    else:
                        return {'status': False, 'message': phd.getstatus()}
                else:
                    return {'status': False,
                            'message': "Keinen Stern gefunden, noch mal probieren oder Position korrigieren. PHD: " + phd.getstatus()}
            else:
                return {'status': False,
                        'message': "Fehler: StartLoop klappt nicht. PHD: " + phd.getstatus()}
        except Exception as e:
            phd = None
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def phd_guiding_stop(self):
        """
        Route '/phd_guiding_stop' - stopt das Guiding.
        :return: JSON-Daten des Status.
        """
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
    def phd_beenden(self):
        """
        Route '/phd_beenden' - beendet PHD.
        :return:
        """
        try:
            bdsrun("phd_beenden")
            # todo: Validieren
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    # </editor-fold>

    # <editor-fold desc="Astrotortilla Routen">

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def at_start(self):
        """
        Route '/at_start' - startet die Software Astrotortilla.
        :return: JSON-Daten des Status.
        """
        try:
            bdsrun("at_start")
            # todo: Start validate
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    @cherrypy.expose
    @cherrypy.tools.json_out()
    def at_platesolve(self):
        """
        Route 'at_platesolve' - Startet das PlateSolving - ohne Validierung.
        :return: JSON-Daten des Status.
        """
        try:
            bdsrun("at_platesolve")
            # todo: Funktion zur Auswertung von Plate Solve Ergebnissen?
            return {'status': True}
        except Exception as e:
            return {'status': False, 'message': str(e)}

    # </editor-fold>

    pass


def background_eqmod_goto_name(objekt, host="", port="", cmd="", key=""):
    """
    Ermittelt die Position des angegebenen Objektes und berechnet diese
    auf AltAz-Koordinaten der aktuellen Sternwartenposition. Anschließend
    wird die Rückantwort zum Server gegeben.
    :param host: Hostname des Servers
    :param port: Port des Servers
    :param cmd: Rückrouten-Befehl
    :param key: Rückrouten-Schlüssel
    :param objekt: GoTo Objekt Name als NGC-Katalogeintrag.
    :return:
    """
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

        if host != "":
            responseserver(host, port, cmd, key, True)

        return True

    except Exception as e:

        if host != "":
            responseserver(host, port, cmd, key, False)

        return False


def bdsrun(name):
    """
    Führt ein BDS-Script aus.
    :param name:  Name des Scriptes.
    :return:
    """
    bdsrun = current_dir + "\\BDSRun.exe /Script:"
    bdsrun_folder = "\\baramundi\\"
    os.popen(bdsrun + current_dir + bdsrun_folder + name + ".bds /S")


def responseserver(host, port, cmd, key, status=True):
    """
    Sorgt bei Aktionen mit Rückantwort für eine Antwort an den Server.
    :param host: Hostname des Servers
    :param port: Port des Servers
    :param cmd: Return-Befehl, üblicherweise 'return'
    :param key: Einmalig erstellter Schlüssel für diese eine Abfrage.
    :param status: Erfolg- oder Misserfolgsmeldung
    :return:
    """
    print("Inform Sender: ", host, port, cmd, key, status)
    conn = http.client.HTTPConnection(host, port)
    conn.request("GET", "/" + cmd + "/" + key + "/" + str(status))
    r1 = conn.getresponse()
    conn.close()
    print(r1.status, r1.reason)
    pass


def validate_password(realm, username, password):
    """
    Validiert das Kennwort mit dem in der Konfiguration hinterlegtem.
    :param realm: unverwendete Zeichenkette
    :param username: unverwendete Zeichenkette
    :param password: Kennwort
    :return: True/False
    """
    return server_challenge == password


if __name__ == '__main__':

    current_dir = os.path.dirname(os.path.abspath(__file__))

    # BackgroundTaskQueue initialisieren
    bgtask = BackgroundTaskQueue(cherrypy.engine)
    bgtask.subscribe()

    # Config lesen
    Config = configparser.ConfigParser()
    Config.read("./config.cfg")
    server_challenge = Config.get("Settings", "ServerChallenge")

    # WebServer cherrypy konfigurieren und starten
    cherrypy.config.update({'tools.auth_basic.checkpassword': validate_password,
                            'tools.auth_basic.on': True,
                            'tools.auth_basic.realm': "localhost"})

    # Mapping der Route '/' auf TMWServer()
    cherrypy.quickstart(TMWServer(), "/", "server.cfg")
