# -*- coding: utf-8 -*-
"""
Created on Fri Jan 28 15:05:12 2022

@author: Bruger
"""
import string
import time
import sys
import array
import os
from comtypes.client import GetModule
from comtypes.client import CreateObject
# Run GetModule once to generate comtypes.gen.VisaComLib.
if not hasattr(sys, "frozen"):
    GetModule("C:\Program Files (x86)\IVI Foundation\VISA\VisaCom\GlobMgr.dll")
import comtypes.gen.VisaComLib as VisaComLib


def do_command(myScope,command):
    myScope.WriteString(""+ command, True)
    #check_instrument_errors(myScope,command)
# =========================================================
# Send a command and check for errors:
# =========================================================
def do_command_ieee_block(myScope,command, data):
    myScope.WriteIEEEBlock(command, data, True)
    #check_instrument_errors(myScope,command)
# =========================================================
# Send a query, check for errors, return string:
# =========================================================
def do_query_string(myScope,query):
    myScope.WriteString(""+ query, True)
    result = myScope.ReadString()
    #check_instrument_errors(myScope,query)
    return result
# =========================================================
# Send a query, check for errors, return string:
# =========================================================
def do_query_ieee_block_UI1(myScope,query):
    myScope.WriteString(""+ query, True)
    result = myScope.ReadIEEEBlock(VisaComLib.BinaryType_UI1, False, True)
    #check_instrument_errors(myScope,query)
    return result
# =========================================================
# Send a query, check for errors, return string:
# =========================================================
def do_query_ieee_block_I2(myScope,query):
    myScope.WriteString(""+ query, True)
    result = myScope.ReadIEEEBlock(VisaComLib.BinaryType_I2, False, True)
    #check_instrument_errors(myScope,query)
    return result
# =========================================================
# Send a query, check for errors, return values:
# =========================================================
def do_query_number(myScope,query):
    myScope.WriteString(""+ query, True)
    result = myScope.ReadNumber(VisaComLib.ASCIIType_R8, True)
    #check_instrument_errors(myScope,query)
    return result     
     
     
     
# =========================================================
# Send a query, check for errors, return values:
# =========================================================
def do_query_numbers(myScope,query):
    myScope.WriteString(""+ query, True)
    result = myScope.ReadList(VisaComLib.ASCIIType_R8, ",;")
    #check_instrument_errors(myScope,query)
    return result
# =========================================================
# Check for instrument errors:
# =========================================================
def check_instrument_errors(myScope,command):
    while True:
        myScope.WriteString(":SYSTem:ERRor? STRing", True)
        error_string = myScope.ReadString()
        if error_string: # If there is an error string value.
            if error_string.find("0,", 0, 2) == -1: # Not "No error".
                print ("ERROR: "+error_string+", command: "+ command)
                print ("Exited because of error.")
                sys.exit(1)
            else: # "No error"
                break
        else: # :SYSTem:ERRor? STRing should always return string.
            print ("ERROR: :SYSTem:ERRor? STRing returned nothing, command: "+command)
            print ("Exited because of error.")
            sys.exit(1)


def initialize(myScope):
# Get and display the device's *IDN? string.
    idn_string = do_query_string(myScope,"*IDN?")
    print ("Identification string "+ idn_string)

    # Clear status and load the default setup.
    do_command(myScope,"*CLS")
    do_command(myScope,"*RST")
# =========================================================
# Capture:
# =========================================================



def setTrigger(myScope,channelNumber,level):
    
    # Set trigger mode. Always assume Edge
    do_command(myScope,":TRIGger:MODE EDGE")
    qresult = do_query_string(myScope,":TRIGger:MODE?")
    print ("Trigger mode: "  +  qresult)
    # Set EDGE trigger parameters.
    do_command(myScope,":TRIGger:EDGE:SOURCe CHANnel{}".format(channelNumber))
    qresult = do_query_string(myScope,":TRIGger:EDGE:SOURce?")
    print( "Trigger edge source: "+ qresult)
    do_command(myScope,":TRIGger:LEVel CHANnel{},{}".format(channelNumber,level))
    qresult = do_query_string(myScope,":TRIGger:LEVel? CHANnel{}".format(channelNumber))
    print ("Trigger level, channel {}: ".format(channelNumber)+ qresult)
    do_command(myScope,":TRIGger:EDGE:SLOPe POSitive")
    qresult = do_query_string(myScope,":TRIGger:EDGE:SLOPe?")
    print ("Trigger edge slope: "+ qresult)    

def setScale(myScope,channelNumber,axis):
    xoffset,timebase,yoffset,vertscale=axis
    
    vertscale=min(1.0,vertscale)
    vertscale=max(1.0/1000,vertscale)
    
    
    # Set vertical scale and offset.
    do_command(myScope,":CHANnel{}:SCALe {}".format(channelNumber,vertscale))
    qresult = do_query_number(myScope,":CHANnel{}:SCALe?".format(channelNumber))
    print ("Channel {} vertical scale: ".format(channelNumber)+ str(qresult))
    do_command(myScope,":CHANnel{}:OFFSet {}".format(channelNumber,str(yoffset)))
    qresult = do_query_number(myScope,":CHANnel{}:OFFSet?".format(channelNumber))
    print ("Channel {} offset: ".format(channelNumber)+ str(qresult))
    
    
    # Set horizontal scale andx offset.
    do_command(myScope,":TIMebase:SCALe {}".format(timebase))
    qresult = do_query_string(myScope,":TIMebase:SCALe?")
    print ("Timebase scale: "+ qresult)
    do_command(myScope,":TIMebase:POSition {}".format(xoffset))
    qresult = do_query_string(myScope,":TIMebase:POSition?")
    print ("Timebase position:"+ qresult)

def displayNone(myScope):
    print("Turning OFF display of alVl channels")
    for i in range(1,5):
        displayChannel(myScope,i,0)
        

def displayAll(myScope):
    print("Turning ON display of all channels")
    for i in range(1,5):
        displayChannel(myScope,i,1)


def displayChannel(myScope,channelNumber,value):
    #Turn the display of the specified channel on or off
    
    if value==1:
        do_command(myScope,":CHANnel{}:DISPlay {}".format(channelNumber,value))
        #qresult = do_query_string(myScope,":CHANnel{}:DISPlay?".format(channelNumber))
        print("Display of channel {} turned ON!".format(channelNumber))
    elif value==0:
        do_command(myScope,":CHANnel{}:DISPlay {}".format(channelNumber,value))
        #qresult = do_query_string(myScope,":CHANnel{}:DISPlay?".format(channelNumber))
        print("Display of channel {} turned OFF!".format(channelNumber))
    else:
        print("Error. Specify a value of 0 or 1 to turn a channel on/off")
        
            
    
    


def getTrace(myScope,channelNumber):
    do_command(myScope,":MEASure:SOURce CHANnel{}".format(channelNumber))
    qresult = do_query_string(myScope,":MEASure:SOURce?")
    print ("Measure source: "+ qresult)
    
    
def getScreenshot(myScope,savePath):
    image_bytes = do_query_ieee_block_UI1(myScope,":DISPlay:DATA? PNG")
    f = open(savePath+".png", "wb")
    f.write(bytearray(image_bytes))
    f.close()
    print ("Screen image written to {}.".format(savePath))    
    


def getWaveform(myScope,channelNumber,savePath):
    
    qresult = do_query_string(myScope,":WAVeform:POINts?")
    numPoints=int(qresult)
    print ("Waveform points: {}".format( numPoints))
    
    # Set the waveform source.
    do_command(myScope,":WAVeform:SOURce CHANnel{}".format(channelNumber))
    qresult = do_query_string(myScope,":WAVeform:SOURce?")
    print ("Waveform source: "+ qresult)
    
    # Choose the format of the data returned:
    do_command(myScope,":WAVeform:FORMat WORD")
    print ("Waveform format: "+ do_query_string(myScope,":WAVeform:FORMat?"))

    wav_form_dict = {
    0 : "ASCii",
    1 : "BYTE",
    2 : "WORD",
    3 : "LONG",
    4 : "LONGLONG",
    }
    acq_type_dict = {
    1 : "RAW",
    2 : "AVERage",
    3 : "VHIStogram",
    4 : "HHIStogram",
    6 : "INTerpolate",
    10 : "PDETect",
    }
    acq_mode_dict = {
    0 : "RTIMe",
    1 : "ETIMe",
    3 : "PDETect",
    }
    coupling_dict = {
    0 : "AC",
    1 : "DC",
    2 : "DCFIFTY",
    3 : "LFREJECT",
    }
    units_dict = {
    0 : "UNKNOWN",
    1 : "VOLT",
    2 : "SECOND",
    3 : "CONSTANT",
    4 : "AMP",
    5 : "DECIBEL",
    }


    preamble_string = do_query_string(myScope,":WAVeform:PREamble?")
    
    
    
    (wav_form, acq_type, wfmpts, avgcnt, x_increment, x_origin,
     x_reference, y_increment, y_origin, y_reference, coupling,
    x_display_range, x_display_origin, y_display_range,
    y_display_origin, date, time_list, frame_model, acq_mode,
    completion, x_units, y_units, max_bw_limit, min_bw_limit
    ) = preamble_string.split(",")
    
    
    if not any(fname.endswith('.txt') for fname in os.listdir('.')):
        
        getScreenshot(myScope,"Screenshot_0")
        
        f = open('Info.txt', "w")
        
        f.write("Waveform format: "+  wav_form_dict[int(wav_form)]+'\n')
        f.write("Acquire type: "+ acq_type_dict[int(acq_type)]+'\n')
        f.write("Waveform points desired: "+ wfmpts+'\n')
        f.write("Waveform average count: "+ avgcnt+'\n')
        f.write("Sampling Rate: "+ str(1/float(x_increment)/1e9)+' GHz'+'\n')  
        f.write("Waveform X increment: "+ x_increment+'\n')
        f.write("Waveform X origin: "+ x_origin+'\n')
        f.write("Waveform X reference: "+ x_reference+'\n') # Always 0.
        f.write("Waveform Y increment: "+ y_increment+'\n')
        f.write("Waveform Y origin: "+ y_origin+'\n')
        f.write("Waveform Y reference: "+ y_reference+'\n') # Always 0.
        f.write("Coupling: "+ coupling_dict[int(coupling)]+'\n')
        f.write("Waveform X display range: "+ x_display_range+'\n')
        f.write("Waveform X display origin: "+ x_display_origin+'\n')
        f.write("Waveform Y display range: "+ y_display_range+'\n')
        f.write("Waveform Y display origin: "+ y_display_origin+'\n')
        f.write("Date: "+ date+'\n')
        f.write("Time: "+ time_list+'\n')
        f.write("Frame model #: "+ frame_model+'\n')
        f.write("Acquire mode: "+ acq_mode_dict[int(acq_mode)]+'\n')
        f.write("Completion pct: "+ completion+'\n')
        f.write("Waveform X units: "+ units_dict[int(x_units)]+'\n')
        f.write("Waveform Y units: "+ units_dict[int(y_units)]+'\n')
        f.write("Max BW limit: "+ max_bw_limit+'\n')
        f.write("Min BW limit: "+ min_bw_limit+'\n')   
         
        
        f.close()
    
    # Get numeric values for later calculations.
    x_increment = do_query_number(myScope,":WAVeform:XINCrement?")
    x_origin = do_query_number(myScope,":WAVeform:XORigin?")
    y_increment = do_query_number(myScope,":WAVeform:YINCrement?")
    y_origin = do_query_number(myScope,":WAVeform:YORigin?")

    do_command(myScope,":WAVeform:STReaming OFF")
    data_words = do_query_ieee_block_I2(myScope,":WAVeform:DATA?")
    nLength = len(data_words)
    print ("Number of data values: "+ str(nLength))
    
    f = open(savePath, "w")
    # Output waveform data in CSV format.
    for i in range(0, nLength):
        
        time_val = x_origin + (i * x_increment)
        voltage = (data_words[i] * y_increment) + y_origin
        f.write("%E, %f\n" % (time_val, voltage))
    # Close output file.
    f.close()
    print ("Waveform format WORD data written to "+ savePath)

def getAverageWaveform(myScope, channelNumber,avgCount,Ntraces, savePath):
    
    do_command(myScope,":ACQUIRE:AVERAGE OFF")
    if avgCount >1:
        
        do_command(myScope,":ACQUIRE:AVERAGE ON")
        
        do_command(myScope,":ACQuire:COUNt 2")
        do_command(myScope,":ACQuire:COUNt {}".format(avgCount))
        do_command(myScope,":ACQUIRE:COMPLETE 100")
 
    
    
    Ndigits=len(str(Ntraces))
    for i in range(0,Ntraces):
        currentSavePath=savePath+'_'+'{:d}'.format(i).zfill(Ndigits)+'.csv'
        print(currentSavePath)
        getWaveform(myScope,channelNumber,currentSavePath)
        
    do_command(myScope,":ACQUIRE:AVERAGE OFF")
    print("Finished taking averages")
    
    
    

