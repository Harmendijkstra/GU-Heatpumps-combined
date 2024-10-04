# -*- coding: utf-8 -*-
"""
Created on Mon Jul  1 19:50:07 2024

@author: LENLUI
"""
# Supposedly, Big-endian (biggest chunk first), and 4x 16bits register.
# Each register consists of 2 bytes, order 2-1 4-3 6-5 8-7 so reversed.
# HOWEVER, according to ModbusPoll it's "Little Endian byte-swapped"

import struct


def reverse_bits_16bit(n):
    # Reverse the bits of a 16-bit integer
    reversed_n = 0
    for i in range(16):
        reversed_n <<= 1
        reversed_n |= (n >> i) & 1
    return reversed_n

def combine_to_64bit(int_list):
    # Combine 4 16-bit integers into a single 64-bit integer
    combined = 0
    for i, val in enumerate(int_list):
        combined |= (val << (16 * i))
    return combined

def reconstruct_float(int_list):
    # Step 1: Reverse the bit order of each 16-bit integer
    # reversed_ints = [reverse_bits_16bit(n) for n in int_list]
    
    # Step 2: Combine the reversed 16-bit integers into a 64-bit integer
    # combined_64bit = combine_to_64bit(reversed_ints)
    combined_64bit = combine_to_64bit(int_list)
    
    # Step 3: Interpret the 64-bit integer as a double-precision float
    float_value = struct.unpack('d', struct.pack('Q', combined_64bit))[0]
    
    return float_value


# CAN BE DONE MUCH MORE SIMPLE:

def reconstruct_float_simple(int_list):
    # Step 1: Pack each 16-bit integer into a 2-byte sequence
    byte_sequence = b''.join(struct.pack('H', n) for n in int_list)
    
    # Step 2: Interpret the 8-byte sequence as a double-precision float
    float_value = struct.unpack('d', byte_sequence)[0]
    
    return float_value



# Example usage

# New: [0, 8192, 4032, 16418, 0, 24576, 8843, 16433, 0, 0, 0, 0, 0, 24576, 3887, 16418, 0, 32768, 12823, 16394, 19833, 8350, 42411, 16546]

# vars = {'Press1':[0, 57344, 42172, 16417], 'Temp1':[0, 24576, 29219, 16437], 'Flow1':[50556, 62587, 23335, 16547],
#         'Press2':[0, 40960, 41732, 16417], 'Temp2':[0, 40960, 5616, 16434], 'Flow2':[0, 0, 0, 0]}

# vars = {'Press1':[0, 8192, 4032, 16418], 'Temp1':[0, 24576, 8843, 16433], 'Flow1':[0, 0, 0, 0],
#         'Press2':[0, 24576, 3887, 16418], 'Temp2':[0, 32768, 12823, 16394], 'Flow2':[19833, 8350, 42411, 16546]}


# Above values, OLD, are tested in 'Modbus Slave' and are: 
    # '64bit Double - Little Endian byte swap
# Below values, NEW from 03-sept, are tested also and are no more byte swap:
    # '64bit Double - Little Endian    
# After 03-sept-2024 EVHI Firmware update:
    # EVHI:[0, 192, 23011, 8512, 0, 160, 58844, 13376, 0, 0, 0, 0, 0, 128, 56546, 8512, 0, 224, 8051, 13376, 10685, 36088, 53651, 42048]

# EVHI = [0, 192, 23011, 8512, 0, 160, 58844, 13376, 0, 0, 0, 0, 0, 128, 56546, 8512, 0, 224, 8051, 13376, 10685, 36088, 53651, 42048]
# EVHI = [0, 160, 19190, 8512, 0, 96, 40968, 13632, 0, 0, 0, 0, 0, 32, 24053, 8512, 0, 160, 12381, 12608, 21220, 14136, 6222, 42304]
EVHI = [0, 64, 60681, 8768, 0, 0, 44643, 13632, 0, 0, 0, 0, 0, 32, 51208, 8768, 0, 96, 7155, 11840, 56436, 11832, 30736, 40000]

# vars = {'Press1':[0, 192, 23011, 8512], 'Temp1':[0, 160, 58844, 13376], 'Flow1':[0, 0, 0, 0],
#         'Press2':[0, 128, 56546, 8512], 'Temp2':[0, 224, 8051, 13376], 'Flow2':[10685, 36088, 53651, 42048]}
vars = {'Press1':EVHI[0:4], 'Temp1':EVHI[4:8], 'Flow1':EVHI[8:12],
        'Press2':EVHI[12:16], 'Temp2':EVHI[16:20], 'Flow2':EVHI[20:24]}


testEVHI01 = [0, 57344, 20342, 16394] #temp str2 2.5degc?
testEVHI01a = [0, 16384, 4131, 16433] #temp str2 normal?
byte_sequence = b''.join(struct.pack('H', n) for n in testEVHI01a)
float_value = round(struct.unpack('d', byte_sequence)[0],3)
print(f'\nValue of TEST value: {float_value}\n')

for signal in vars:
    int_list = vars[signal]  # Replace with your actual 16-bit integers
    # int_list_reversed = [reverse_bits_16bit(n) for n in int_list]
    # float_value = round(reconstruct_float(int_list),4)
    
    # Step 1: Pack each 16-bit integer into a 2-byte sequence
    byte_sequence = b''.join(struct.pack('>H', n) for n in int_list)
    # Step 2: Interpret the 8-byte sequence as a double-precision float
    float_value = round(struct.unpack('<d', byte_sequence)[0],3)
    # float_value = round(reconstruct_float_simple(int_list),4)
    print(f'Value of {signal}: {float_value}')
    
# Now process the little endian byte swapped data from Belimo Flow:
    # 1. Little endian: swap the 2 bytes of each 16bit UINT first
    # 2. LSW or Least Significant Word first, so a register swap: interpret 2 registers as low-high word.
# first 2 UNITS are flow in m3 (divide by 100), 2nd set is flow in gallons (exact), 3rd set is kWh totalized
# NOTE: uint16 values are well received below, no need for a bit reversal. Just a register swap.
belimoFlow = [29853, 0, 13327, 1, 299, 0]
belimo01 = [33631, 0, 23308, 1, 336, 0]
belimo02 = [12692, 0, 33527, 0, 127, 0]
belimo03 = [27967, 0, 8346, 1, 280, 0]
belimoValve = [2989, 0, 7896, 0, 30, 0]
belimoUnits = ['m3*100','gallons','m3']

belimoFlow = belimoValve

belimo01instFlow = [ 4646, 8]

print('\n')
belimoFlowBitsReversed = []
for i in range(0,3):
    # belimoFlowBitsReversed.extend([reverse_bits_16bit(x) for x in belimoFlow[i*2:i*2+2]])
    # byte_sequence = b''.join(struct.pack('H', n) for n in belimoFlowBitsReversed[i*2:i*2+2])
    byte_sequence = b''.join(struct.pack('<H', n) for n in belimoFlow[i*2:i*2+2])
    uint32value = struct.unpack('<I', byte_sequence)[0]
    print(f'for i=={i} we get {uint32value} {belimoUnits[i]}')
    
# belimoWrongUINT32 = 21999780
# belimoUINTs16 = [335, 27660]
# byte_sequence = struct.pack('>I', belimoWrongUINT32)
# struct.unpack('<I', byte_sequence)[0]

# belimoUINTs16 = [27660, 335]
# byte_sequence = b''.join(struct.pack('H', n) for n in belimoUINTs16)


# Wrongly interpreted as big-endian UINT32 AND then multiplied by 0.01:
# belimoWrongUINT32 = 21999780*100 #21982080 
belimoWrongUINT32 = 2199977984 # Precision was lost during logging
res = struct.pack('>I', belimoWrongUINT32)
corr = struct.unpack(">HH", res)
    # Out[263]: (33569,)
print(f'\nWith wrong UINT32 Big Endian {belimoWrongUINT32} we get {corr}')

# Convert that value back into this:
# BelimoUINTs16correct = [33627, 0]
BelimoUINTs16correct = [33550, 0]
byte_seq_correct = b''.join(struct.pack('>H', n) for n in BelimoUINTs16correct)
wrongres = struct.unpack(">I", byte_seq_correct)[0]
    # Out[282]: (2198732800,)
# 10-figure output
print(f'With input {BelimoUINTs16correct} we get wrong UINT32 Big Endian {wrongres}')



