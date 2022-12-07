# -*- coding: utf-8 -*-
#this module is to write the control functions of SMU 2611B
# dated:9/17/2022
# author:Jiawen Zhou
import serial
import time
import pyvisa

rm=pyvisa.ResourceManager()
#Use usb connection here

def open_Keithley2611B(obj,Keithley2611B_port=''):
    '''
    Open Keithley 2611B connection, use VISA usb port
    :param obj: GUI Class obj
    :param mpc_port: Keithley2611B port number，maynot need, will auto detect and try
    :return: True/False,keithley port detected
    '''
    try:
        equip_list=rm.list_resources()
        keithley=[i for i in equip_list if 'USB' in i]
        Keithley2611B_port='NA'
        if keithley==[]:
            print('Keithley SMU not found, please check USB connection!')
            return False,'NA'
        for i in keithley:
            if not i in rm.list_opened_resources():
                obj.SMU_2611B=rm.open_resource(i)
                obj.SMU_2611B.timeout=200
                ff=obj.SMU_2611B.query('*IDN?')
                if 'Keithley Instruments Inc' in ff:
                    print(ff)
                    Keithley2611B_port=i
                    return True,Keithley2611B_port
                else:
                    print('This is not 2611B:',i)
                if keithley.index(i)==len(keithley)-1:
                    print('未问询到2611B,请检查连接！')
                    return False,'NA'
            else:
                print('Port opened, go to the next:',i)

    except Exception as e:
        print(e)
        return False,Keithley2611B_port

def open_Keithley2611B_unused(obj,Keithley2611B_port):
    '''
    Open Keithley 2611B connection, use VISA port or Serial port
    :param obj: GUI Class obj
    :param mpc_port: MPC port number
    :return: True/False
    '''
    try:
        if 'INSTR' in str(Keithley2611B_port).upper():
            print('Keithley2611B VISA port connection...')
            obj.SMU = rm.open_resource(Keithley2611B_port)
            obj.SMU.timeout = 5
            ins = obj.SMU.query('*IDN?')
            if "Keithley2611B" in ins:
                return True
            else:
                return False
        elif 'COM' in str(Keithley2611B_port).upper():
            print('Keithley2611B serial port connection...')
            obj.SMU = serial.Serial(Keithley2611B_port,115200,timeout=5)
            return True
    except Exception as e:
        print(e)
        return False

def close_Keithley2611B(obj):
    '''
    Close Keithley2611B connection
    :param obj: GUI Class obj:
    :return: True/False
    '''
    try:
        if Keithley2611B_port in rm.list_opened_resources():
            obj.SMU_2611B.close()
            print('Keithley2611B通讯关闭成功!')
        else:
            print('Keithley2611B通讯未打开!')
    except Exception as e:
        print(e)
        return

def Initiate_Keithley2611B_alignment(obj,Voutput=1,Currentlimit=0.001):
    '''
    Initiate alignment settings of SMU
    :param obj:
    :param Voutput: output setting, omit is 1V
    :param currentlimit:  here for alignment is 1mA
    :return:
    '''
    #firstly set SMU output to 1V and voltage mode
    Keithley2612B_SetSource_VoltageMode(obj)
    Keithley2612B_SetSource_VoltageLevel(obj,Voutput)
    Keithley2612B_SetSource_CurrentLimit(obj,Currentlimit)
    resu=Keithley2612B_MeasureCurrent(obj)
    return resu

def Keithley2612B_SetSource_CurrentLimit(obj,current):
    '''
    Function to set limit current
    :param obj:
    :param current:
    :return:
    '''
    command="smua.source.limiti = " + str(current) + "\n"
    obj.SMU_2611B.write(command)
    time.sleep(0.1)

def Keithley2612B_SetSource_VoltageMode(obj):
    '''
    Function that sets the voltage source mode of the source meter
    :param obj: 
    :param voltage: 
    :return: 
    '''
    command = "smua.source.func = smua.OUTPUT_DCVOLTS\n"
    obj.SMU_2611B.write(command)
    time.sleep(0.1)

def Keithley2612B_SetSource_VoltageLevel(obj,voltage):
    '''
    Function that sets the voltage source mode of the source meter
    :param obj:
    :param voltage:
    :return:
    '''
    command = "smua.source.levelv = " + str(voltage) + "\n"
    obj.SMU_2611B.write(command)
    time.sleep(0.1)

def Keithley2612B_Enable_Sense_Current_AutoRange(obj):
    '''
    Function to enable sense auto range.
    :param obj:
    :param voltage:
    :return:
    '''
    command = "smua.measure.autorangei = smua.AUTORANGE_ON\n"
    obj.SMU_2611B.write(command)
    time.sleep(0.1)

def Keithley2612B_MeasureCurrent(obj):
    '''
    Function to measure current.
    :return:current measured
    '''
    command = "READING = smua.measure.i()\n"
    obj.SMU_2611B.write(command)
    time.sleep(0.1)

    command = "print(READING)\n"
    obj.SMU_2611B.write(command)
    time.sleep(0.1)

    respon=float(obj.SMU_2611B.read())

    # command = "display.settext(READING..\"$N\")\n"
    # obj.SMU_2611B.write(command)
    # time.sleep(0.1)
    return respon

def openOutput_Keithley2611B(obj):
    '''
    Open Keithley2611B output
    :param obj: GUI Class obj:
    :return: NA
    '''
    command = "smua.source.output = smua.OUTPUT_ON\n"
    obj.SMU_2611B.write(command)
    time.sleep(0.1)


def CloseOutput_Keithley2611B(obj):
    '''
    Close Keithley2611B output
    :param obj: GUI Class obj:
    :return: NA
    '''
    command = "smua.source.output = smua.OUTPUT_OFF\n"
    obj.SMU_2611B.write(command)
    time.sleep(0.1)

def reset_Keithley2611B(obj):
    '''
    Open Keithley2611B output
    :param obj: GUI Class obj:
    :return: NA
    '''
    command = "*RST\n"
    obj.SMU_2611B.write(command)
    time.sleep(0.1)



'''
/// Functions to control the source meter Keithley 2612B through VISA.

/// Function to initialize variables used by the Keithley2612B module.
/// Input: -
/// Output: $KeithleyResourceName
///
function Keithley2612B_Init()
{
  $KeithleyResourceName = "GPIB0::26::INSTR";
  
  return;
}

/// Function that sets the voltage source mode of the source meter.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_SetSource_VoltageMode()
{
  $Command = "smua.source.func = smua.OUTPUT_DCVOLTS\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to set the source voltage level.
/// Input: $Keithley2612B, $Voltage
/// Output: -
///
function Keithley2612B_SetSource_VoltageLevel()
{
  $Command = "smua.source.levelv = " + $Voltage + "\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to set the source current limit.
/// Input: $Keithley2612B, $Current
/// Output: -
///
function Keithley2612B_SetSource_CurrentLimit()
{
  $Command = "smua.source.limiti = " + $Current + "\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to set the sense current range.
/// Input: $Keithley2612B, $Current
/// Output: -
///
function Keithley2612B_SetSense_CurrentRange()
{
  $Command = "smua.measure.rangei = " + $Current + "\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to enable sense auto range.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_Enable_Sense_Current_AutoRange()
{
  $Command = "smua.measure.autorangei = smua.AUTORANGE_ON\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to disable sense auto range.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_Disable_Sense_Current_AutoRange()
{
  $Command = "smua.measure.autorangei = smua.AUTORANGE_OFF\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to enable sense auto zero.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_Enable_Sense_AutoZero()
{
  $Command = "smua.measure.autozero = smua.AUTOZERO_AUTO\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to disable sense auto zero.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_Disable_Sense_AutoZero()
{
  $Command = "smua.measure.autozero = smua.AUTOZERO_OFF\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to sense auto zero once.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_Sense_AutoZero_Once()
{
  $Command = "smua.measure.autozero = smua.AUTOZERO_ONCE\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to turn the output on.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_OutputOn()
{
  $Command = "smua.source.output = smua.OUTPUT_ON\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to turn the output off.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_OutputOff()
{
  $Command = "smua.source.output = smua.OUTPUT_OFF\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to take current measurements.
/// Input: $Keithley2612B,
/// Output: $Current
///
function Keithley2612B_MeasureCurrent()
{
  $Command = "READING = smua.measure.i()\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  $Command = "print(READING)\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  $Response = IviVISA_Read($Keithley2612B);
  $Current = StringParseToFloat($Response);
  
  $Command = "display.settext(READING..\"$N\")\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to set up a linear sweep specifying the number of points between start and stop to get the measurements.
/// Input: $Keithley2612B, $Start, $Stop, $Points, $Delay
/// Output: -
///
function Keithley2612B_SetUpLinearVoltageSweepByPoints()
{
  $Command = "smua.nvbuffer1.collectsourcevalues = 1\n"; /// enable collection of source values during sweeps
  IviVISA_Write($Keithley2612B,$Command);
  $Command = "SweepVLinMeasureI(smua, " + $Start + ", " + $Stop + ", " + $Delay + ", " + $Points + ")\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to clear the data buffer to store the sweep data.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_ClearDataBuffer()
{
  $Command = "smua.nvbuffer1.clear()\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to retrieve data from the source meter buffer.
/// Input: $Keithley2612B
/// Output: $SenseData, $SourceData, $TraceDataWasRead
///
function Keithley2612B_ReadDataBuffer()
{
  $SenseData = "";
  $TraceDataWasRead = false;
  
  $Command = "printbuffer(1, " + $Points + ", smua.nvbuffer1.readings)\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(200);
  
  ($Response,$ReadSuccess) = IviVISA_ReadBool($Keithley2612B);
  if ($ReadSuccess)
  {
    $SenseData = $Response;
    $Command = "printbuffer(1, " + $Points + ", smua.nvbuffer1.sourcevalues)\n";
    IviVISA_Write($Keithley2612B,$Command);
    SleepMilliseconds(200);
    
    ($Response,$ReadSuccess) = IviVISA_ReadBool($Keithley2612B);
    if ($ReadSuccess)
    {
      $SourceData = $Response;
      $TraceDataWasRead = true;
    }
  }
  
  return;
}

/// Function to parse the trace data and put it in data arrays.
/// Input: $Keithley2612B, $SenseData, $SourceData, $Points
/// Output: $MeasurementArray, $SourceArray
///
function Keithley2612B_ParseTraceData()
{
  $MeasurementArray = Array1DCreate("FLOAT",$Points);
  $SourceArray = Array1DCreate("FLOAT",$Points);
  
  $SenseStringArray = StringSplitToArray($SenseData,",",$Points);
  $SourceStringArray = StringSplitToArray($SourceData,",",$Points);
  
  $Index = 0;
  while ($Index < $Points)
  {
    $StringValue = Array1DGetValue($SenseStringArray,$Index);
    $Value = StringParseToFloat($StringValue);
    Array1DSetValue($MeasurementArray,$Index,$Value);
    
    $StringValue = Array1DGetValue($SourceStringArray,$Index);
    $Value = StringParseToFloat($StringValue);
    Array1DSetValue($SourceArray,$Index,$Value);
    
    $Index = $Index + 1;
  }
  
  return;
}

/// Function to configure display.
/// Input: $Keithley2612B,
/// Output: -
///
function Keithley2612B_ResetDisplay()
{
  /// Display only channel A.
  $Command = "display.screen = display.SMUA\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  /// Display current measurements.
  $Command = "display.smua.measure.func = display.MEASURE_DCAMPS\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to clear the display.
/// Input: $Keithley2612B,
/// Output: -
///
function Keithley2612B_ClearDisplay()
{
  /// Display only channel A.
  $Command = "display.clear()\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  return;
}

/// Function to check if there is any pending operation running on the source meter. It will return 1 when all pending 
/// operations complete, or will time out and fail otherwise.
/// Input: $Keithley2612B,
/// Output: $Status
///
function Keithley2612B_OperationComplete()
{
  $Command = "*OPC?\n";
  IviVISA_Write($Keithley2612B,$Command);
  SleepMilliseconds(100);
  
  $Response = IviVISA_Read($Keithley2612B);
  $Status = $Response;
  
  return;
}

/// Function to open the handle to communicate with the source meter.
/// Input: $KeithleyResourceName
/// Output: $Keithley2612B
///
function Keithley2612B_Open()
{
  $Exists = CheckVariableExists("Keithley2612B_Timeout");
  if (!$Exists)
  {
    $Keithley2612B_Timeout = 60000;
  }
  $Keithley2612B = IviVISA_Open($KeithleyResourceName,$Keithley2612B_Timeout);  /// timeout must be longer than the longest sweep
  IviVISA_Write($Keithley2612B,"*IDN?\n");
  $Response = IviVISA_Read($Keithley2612B);
  UpdateStatus($Response);
  
  return;
}

/// Function that resets the source meter.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_Reset()
{
  IviVISA_Write($Keithley2612B,"*RST\n");
  Sleep(1);
  
  return;
}

/// Function to close the communication with the source meter.
/// Input: $Keithley2612B
/// Output: -
///
function Keithley2612B_Close()
{
  IviVISA_Close($Keithley2612B);
  
  return;
}

'''