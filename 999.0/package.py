# -*- coding: utf-8 -*-

name = "start_multi_app"
version = "999.0"
description = "Program launcher with sorted process controls"
authors = ["Lugwit Team"]

requires = [
    "python-3.12+<3.13",
    "Lugwit_Module",
]

build_command = False
cachable = True
relocatable = True


def commands():
    env.PYTHONPATH.prepend("{root}/src")
    alias("start_multi_app", "python {root}/src/start_multi_app/main.py")

