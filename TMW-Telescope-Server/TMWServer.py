# -*- coding: utf-8 -*-

import time
import os
import io
import cherrypy
import configparser
import datetime

import subprocess

import win32com
import win32com.client
import pythoncom
from cherrypy.lib import auth_basic, file_generator


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
    def eqmod_start(self):
        def eqmod_starten():
            try:
                yield "EQMod - Starten...<br>"
                #os.popen(self.current_dir + "\\eqascom_startup.vbs")
                #yield "Status pr&uml;fen...<br>"
                #time.sleep(2)

                pythoncom.CoInitialize()

                o = win32com.client.Dispatch("EQMOD.Telescope")
                o.Connected = True
                o.IncClientCount()

                if o.CanSlew:
                    yield "EQMod - erfolgreich gestartet" + "<br>"
                else:
                    yield "EQMod - Fehler beim Starten von EQMod (Teleskop verbunden & eingeschaltet?)." + "<br>"

            except Exception as e:
                yield "EQMod - Fehler beim Starten von EQMod." + "<br>" + e.message + "<br>"

        return eqmod_starten()


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

