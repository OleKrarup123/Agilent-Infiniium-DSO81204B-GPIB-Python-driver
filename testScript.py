# -*- coding: utf-8 -*-
"""
Created on Sun Jan 30 15:09:57 2022

@author: Bruger
"""
# -*- coding: utf-8 -*-
"""
Created on Fri Jan 28 13:31:08 2022

@author: Bruger
"""
#
# Agilent VISA COM Example in Python using "comtypes"
# *********************************************************
# This program illustrates a few commonly used programming
# features of your Agilent Infiniium Series oscilloscope.
# *********************************************************
# Import Python modules.
# ---------------------------------------------------------

import time
import sys
import os
import inspect
from comtypes.client import GetModule
from comtypes.client import CreateObject
from datetime import datetime 
# Run GetModule once to generate comtypes.gen.VisaComLib.
if not hasattr(sys, "frozen"):
    GetModule("C:\Program Files (x86)\IVI Foundation\VISA\VisaCom\GlobMgr.dll")
import comtypes.gen.VisaComLib as VisaComLib

from OSCfunctions import *
import pyvisa as pv #Library for sending GPIB commands via python

     
rm = CreateObject("VISA.GlobalRM", interface=VisaComLib.IResourceManager)
myScope = CreateObject("VISA.BasicFormattedIO", interface=VisaComLib.IFormattedIO488)


GPIBlist=pv.ResourceManager().list_resources() #Get list of all devices connected via GPIB 
print(GPIBlist)

GPIBname=GPIBlist[0] #Pick the right device name
print(GPIBname)

myScope.IO = rm.Open(GPIBname)
# Clear the interface.
myScope.IO.Clear()
print ("Interface cleared.")
# Set the Timeout to 15 seconds.
myScope.IO.Timeout = 15000 # 15 seconds.
print ("Timeout set to 15000 milliseconds.")
     
         
channelNumber=4 #Set channel to measure and trigger on
triggerLevel=1.5 #Set trigger level


initialize(myScope)  #Check connection to scope
displayNone(myScope) #Turn all traces off
displayChannel(myScope,channelNumber,1) #Turn specified channel on


setTrigger(myScope,channelNumber,triggerLevel) #Set trigger channel and level


#Specify axis of trace
xoffset, timebase, yoffset, vertscale = 0, 100e-9,  0,  1
axis=[xoffset,timebase,yoffset,vertscale]
setScale(myScope,4,axis)



avgCount=52 #Number of traces to average for a measurement
Ntraces=12   #Number of measurements

#Generate new folder to store data from this run
now = datetime.now()
dt_string=now.strftime("%Y%m%d%H%M%S")
folder_name='TestExperiment'+'_'+dt_string
savename='TestTrace'

filename = inspect.getframeinfo(inspect.currentframe()).filename
path = os.path.dirname(os.path.abspath(filename))+'\\DataFolder\\'

os.chdir(path)
os.mkdir(folder_name)
os.chdir(folder_name)

savePath=path+folder_name+'\\'+savename


# Run script for acquiring data
getAverageWaveform(myScope, channelNumber,avgCount,Ntraces, savePath)
 
   
print( "End of program")
             