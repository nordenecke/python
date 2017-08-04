# -*- coding: utf-8 -*-
"""
Created on Fri Jul 28 21:46:02 2017

@author: norden
"""

import pip
from subprocess import call

for dist in pip.get_installed_distributions():
    call("pip install --upgrade " + dist.project_name, shell=True)