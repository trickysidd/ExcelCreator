[app]
title = HelloKivy
package.name = hellokivy
package.domain = org.example
source.dir = .
source.include_exts = py,kv,png,jpg,ttf
version = 1.0.0
requirements = python3,kivy
orientation = portrait
fullscreen = 1
android.api = 30
android.minapi = 21
android.ndk = 23b
android.ndk_api = 21
android.arch = armeabi-v7a
android.permissions = INTERNET
android.allow_backup = False
android.logcat_filters = *:S python:D
android.entrypoint = main.py

[buildozer]
log_level = 2
warn_on_root = 1
