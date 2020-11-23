# -*- coding: utf-8 -*-
"""
Created on Fri Oct  9 12:31:40 2020

@author: huang.zq
"""

import tkinter as tk
import tkinter.filedialog
import xlrd
import xlsxwriter
import pandas as pd
from multiprocessing import Array, Value
from multiprocessing import sharedctypes
import time

#### 查询库容函数
def look_SKvol(height):
    
    if type(height)== sharedctypes.Synchronized:
        height = height.value
    else:
        pass

    if height - SK_hgts[0] < -0.1 or height - SK_hgts[-1] > 0.1:
        raise Exception("ERROR:高程值超过范围！")
    else:
        SK_hgts_pd = pd.Series(list(SK_hgts)).append(pd.Series(height,index=["interest"]))
        height_rank = int(SK_hgts_pd.rank()["interest"]-1)
        volumn = SK_vols[height_rank - 1] + (height-SK_hgts[height_rank - 1])*(SK_vols[height_rank] - SK_vols[height_rank - 1])/(SK_hgts[height_rank] - SK_hgts[height_rank-1])
    return volumn

def look_XKvol(height):
    
    if type(height)== sharedctypes.Synchronized:
        height = height.value
    else:
        pass

    if height - XK_hgts[0] < -0.1 or height - XK_hgts[-1] > 0.1:
        raise Exception("ERROR:高程值超过范围！")
    else:
        XK_hgts_pd = pd.Series(list(XK_hgts)).append(pd.Series(height,index=["interest"]))
        height_rank = int(XK_hgts_pd.rank()["interest"]-1)
        volumn = XK_vols[height_rank - 1] + (height-XK_hgts[height_rank - 1])*(XK_vols[height_rank] - XK_vols[height_rank - 1])/(XK_hgts[height_rank] - XK_hgts[height_rank-1])
    return volumn
#### 查询水位函数
def look_SKhgt(volumn):
    
    if type(volumn)== sharedctypes.Synchronized:
        volumn = volumn.value
    else:
        pass

    if volumn - SK_vols[0] < -0.1 or volumn - SK_vols[-1] > 0.1:
        raise Exception("ERROR:库容值超过范围！")
    else:
        SK_vols_pd = pd.Series(list(SK_vols)).append(pd.Series(volumn,index=["interest"]))
        volumn_rank = int(SK_vols_pd.rank()["interest"]-1)
        height = SK_hgts[volumn_rank - 1] + (volumn-SK_vols[volumn_rank - 1])*(SK_hgts[volumn_rank] - SK_hgts[volumn_rank - 1])/(SK_vols[volumn_rank] - SK_vols[volumn_rank-1])
    return height

def look_XKhgt(volumn):
    
    if type(volumn)== sharedctypes.Synchronized:
        volumn = volumn.value
    else:
        pass

    if volumn - XK_vols[0] < -0.1 or volumn - XK_vols[-1] > 0.1:
        raise Exception("ERROR:库容值超过范围！")
    else:
        XK_vols_pd = pd.Series(list(XK_vols)).append(pd.Series(volumn,index=["interest"]))
        volumn_rank = int(XK_vols_pd.rank()["interest"]-1)
        height = XK_hgts[volumn_rank - 1] + (volumn-XK_vols[volumn_rank - 1])*(XK_hgts[volumn_rank] - XK_hgts[volumn_rank - 1])/(XK_vols[volumn_rank] - XK_vols[volumn_rank-1])
    return height


### 二分法估算调节库容，并进而估算上库正常蓄水位
def adjust_vol_try_func(SK_dth_hgt,XK_dth_hgt):
    
    if type(SK_dth_hgt)== sharedctypes.Synchronized:
        SK_dth_hgt = SK_dth_hgt.value
    else:
        pass
    if type(XK_dth_hgt)== sharedctypes.Synchronized:
        XK_dth_hgt = XK_dth_hgt.value
    else:
        pass
    if type(energy_storage)== sharedctypes.Synchronized:
        energy_storage_i = energy_storage.value
    else:
        energy_storage_i = energy_storage
    if type(min_grossloss_div_netloss)== sharedctypes.Synchronized:
        min_grossloss_div_netloss_i = min_grossloss_div_netloss.value
    else:
        min_grossloss_div_netloss_i = min_grossloss_div_netloss
    if type(max_grossloss_div_netloss)== sharedctypes.Synchronized:
        max_grossloss_div_netloss_i = max_grossloss_div_netloss.value
    else:
        max_grossloss_div_netloss_i = max_grossloss_div_netloss
    if type(effic_coef)== sharedctypes.Synchronized:
        effic_coef_i = effic_coef.value
    else:
        effic_coef_i = effic_coef
    
    adjust_vol_try0 = 0
    adjust_vol_try1 = min([SK_vols[-1]-look_SKvol(SK_dth_hgt),XK_vols[-1]-look_XKvol(XK_dth_hgt)])
    adjust_vol_try2 = (adjust_vol_try0 + adjust_vol_try1)/2
    energy_storage_try2 = 0
    
    while abs(energy_storage_i-energy_storage_try2) > energy_storage_i*0.01:
        energy_storage_try2 = adjust_vol_try2 * effic_coef_i * ((look_SKhgt(adjust_vol_try2 + look_SKvol(SK_dth_hgt)) - XK_dth_hgt) +\
                                                             SK_dth_hgt - look_XKhgt(adjust_vol_try2 + look_XKvol(XK_dth_hgt)))/2 *\
                                                             (2-(min_grossloss_div_netloss_i+max_grossloss_div_netloss_i)/2)/3600
        if (energy_storage_try2 > energy_storage_i):
            adjust_vol_try0 = adjust_vol_try0
            adjust_vol_try1 = adjust_vol_try2
            adjust_vol_try2 = (adjust_vol_try0 + adjust_vol_try1)/2
        else:
            adjust_vol_try0 = adjust_vol_try2
            adjust_vol_try1 = adjust_vol_try1
            adjust_vol_try2 = (adjust_vol_try0 + adjust_vol_try1)/2
        print([adjust_vol_try0,adjust_vol_try1,adjust_vol_try2])
    return adjust_vol_try0

### 水头损失计算函数(插值法，粗算)
def loss_func1(gross_head, max_gross_head, min_gross_head):
    
    if type(min_grossloss_div_netloss)== sharedctypes.Synchronized:
        min_grossloss_div_netloss_i = min_grossloss_div_netloss.value
    else:
        min_grossloss_div_netloss_i = min_grossloss_div_netloss
    if type(max_grossloss_div_netloss)== sharedctypes.Synchronized:
        max_grossloss_div_netloss_i = max_grossloss_div_netloss.value
    else:
        max_grossloss_div_netloss_i = max_grossloss_div_netloss
    
    if type(gross_head)== sharedctypes.Synchronized:
        gross_head = gross_head.value
    else:
        pass
    if type(max_gross_head)== sharedctypes.Synchronized:
        max_gross_head = max_gross_head.value
    else:
        pass
    if type(min_gross_head)== sharedctypes.Synchronized:
        min_gross_head = min_gross_head.value
    else:
        pass
    if type(installed_capcity)== sharedctypes.Synchronized:
        installed_capcity_i = installed_capcity.value
    else:
        installed_capcity_i = installed_capcity
    if type(effic_coef)== sharedctypes.Synchronized:
        effic_coef_i = effic_coef.value
    else:
        effic_coef_i = effic_coef
    
    net_head = gross_head / (min_grossloss_div_netloss + (max_gross_head - gross_head)\
                             * (max_grossloss_div_netloss_i - min_grossloss_div_netloss_i)\
                             /(max_gross_head - min_gross_head))
    head_loss = gross_head - net_head
    gener_flow = installed_capcity_i * 10000 / effic_coef_i / net_head
    return(net_head,head_loss,gener_flow)

### 水头损失计算函数(公式法)
def loss_func(gross_head):
    
    if type(gross_head)== sharedctypes.Synchronized:
        gross_head = gross_head.value
    else:
        pass
    if type(gener_loscoef_all)== sharedctypes.Synchronized:
        gener_loscoef_all_i = gener_loscoef_all.value
    else:
        gener_loscoef_all_i = gener_loscoef_all
    if type(installed_nums)== sharedctypes.Synchronized:
        installed_nums_i = installed_nums.value
    else:
        installed_nums_i = installed_nums
    if type(installed_capcity)== sharedctypes.Synchronized:
        installed_capcity_i = installed_capcity.value
    else:
        installed_capcity_i = installed_capcity
    if type(effic_coef)== sharedctypes.Synchronized:
        effic_coef_i = effic_coef.value
    else:
        effic_coef_i = effic_coef
    
    gener_flow1 = installed_capcity_i * 10000 / effic_coef_i / gross_head
    gener_flow2 = 0
    
    while abs(gener_flow1 - gener_flow2) > 0.001:
        head_loss = gener_loscoef_all_i * (gener_flow1 / installed_nums_i)**2 / 1000
        net_head = gross_head - head_loss
        gener_flow2 = installed_capcity_i * 10000 / effic_coef_i / net_head
        gener_flow1 = gener_flow1 + 0.001
        
    head_loss = gener_loscoef_all_i * (gener_flow1 / installed_nums_i)**2 / 1000
    net_head = gross_head - head_loss
    
    return(net_head,head_loss,gener_flow1)
    
### 计算最小净水头（全机组满发）

def min_net_head(XK_norm_hgt, SK_dth_hgt):
    
    if type(XK_norm_hgt)== sharedctypes.Synchronized:
        XK_norm_hgt = XK_norm_hgt.value
    else:
        pass
    if type(SK_dth_hgt)== sharedctypes.Synchronized:
        SK_dth_hgt = SK_dth_hgt.value
    else:
        pass
    if type(installed_nums)== sharedctypes.Synchronized:
        installed_nums_i = installed_nums.value
    else:
        installed_nums_i = installed_nums
    if type(installed_capcity)== sharedctypes.Synchronized:
        installed_capcity_i = installed_capcity.value
    else:
        installed_capcity_i = installed_capcity
    if type(effic_coef)== sharedctypes.Synchronized:
        effic_coef_i = effic_coef.value
    else:
        effic_coef_i = effic_coef
    if type(gener_loscoef_all)== sharedctypes.Synchronized:
        gener_loscoef_all_i = gener_loscoef_all.value
    else:
        gener_loscoef_all_i = gener_loscoef_all

    if decim_plc.value == 1:
        gross_head = SK_dth_hgt - round(XK_norm_hgt+0.05,1)
    else:
        gross_head = SK_dth_hgt - round(XK_norm_hgt+0.5,0)    
        
    gener_flow1 = installed_capcity_i * 10000 / effic_coef_i / gross_head
    gener_flow2 = 0
    
    while abs(gener_flow1 - gener_flow2) > 0.001:
        head_loss = gener_loscoef_all_i * (gener_flow1 / installed_nums_i)**2 / 1000
        net_head = gross_head - head_loss
        gener_flow2 = installed_capcity_i * 10000 / effic_coef_i / net_head
        gener_flow1 = gener_flow1 + 0.001
    
    head_loss = gener_loscoef_all_i * (gener_flow1 / installed_nums_i)**2 / 1000
    net_head = gross_head - head_loss
    gener_flow2 = installed_capcity_i * 10000 / effic_coef_i / net_head
    gener_flow_min = gener_flow2
    net_head_min = net_head
    
    return(gross_head, head_loss, gener_flow_min, net_head_min)
    
### 计算最大净水头（一台机满发）
def max_net_head(SK_norm_hgt, XK_dth_hgt):
    
    if type(SK_norm_hgt)== sharedctypes.Synchronized:
        SK_norm_hgt = SK_norm_hgt.value
    else:
        pass
    if type(XK_dth_hgt)== sharedctypes.Synchronized:
        XK_dth_hgt = XK_dth_hgt.value
    else:
        pass
    if type(installed_capcity_one)== sharedctypes.Synchronized:
        installed_capcity_one_i = installed_capcity_one.value
    else:
        installed_capcity_one_i = installed_capcity_one
    if type(effic_coef)== sharedctypes.Synchronized:
        effic_coef_i = effic_coef.value
    else:
        effic_coef_i = effic_coef
    if type(gener_loscoef_one)== sharedctypes.Synchronized:
        gener_loscoef_one_i = gener_loscoef_one.value
    else:
        gener_loscoef_one_i = gener_loscoef_one
    
    
    gross_head = SK_norm_hgt - XK_dth_hgt
    gener_flow1 = installed_capcity_one_i * 10000 / effic_coef_i / gross_head
    gener_flow2 = 0
    
    while abs(gener_flow1 - gener_flow2) > 0.001:
        head_loss = gener_loscoef_one_i * (gener_flow1 * 0.2 )**2 / 1000
        net_head = gross_head - head_loss
        gener_flow2 = installed_capcity_one_i * 10000 / effic_coef_i / net_head
        gener_flow1 = gener_flow1 + 0.001
    
    head_loss = gener_loscoef_one_i * (gener_flow1 * 0.2 )**2 / 1000
    net_head = gross_head - head_loss
    gener_flow2 = installed_capcity_one_i * 10000 / effic_coef_i / net_head
    gener_flow_max = gener_flow2
    net_head_max = net_head
    
    return(gross_head, head_loss, gener_flow_max, net_head_max)

### 计算平均净水头
def average_net_head(SK_norm_hgt, SK_dth_hgt, XK_norm_hgt, XK_dth_hgt):
    
    if type(SK_norm_hgt)== sharedctypes.Synchronized:
        SK_norm_hgt = SK_norm_hgt.value
    else:
        pass
    if type(XK_dth_hgt)== sharedctypes.Synchronized:
        XK_dth_hgt = XK_dth_hgt.value
    else:
        pass
    if type(XK_norm_hgt)== sharedctypes.Synchronized:
        XK_norm_hgt = XK_norm_hgt.value
    else:
        pass
    if type(SK_dth_hgt)== sharedctypes.Synchronized:
        SK_dth_hgt = SK_dth_hgt.value
    else:
        pass
    
    if type(installed_nums)== sharedctypes.Synchronized:
        installed_nums_i = installed_nums.value
    else:
        installed_nums_i = installed_nums
    if type(installed_capcity)== sharedctypes.Synchronized:
        installed_capcity_i = installed_capcity.value
    else:
        installed_capcity_i = installed_capcity
    if type(installed_capcity_one)== sharedctypes.Synchronized:
        installed_capcity_one_i = installed_capcity_one.value
    else:
        installed_capcity_one_i = installed_capcity_one
    if type(effic_coef)== sharedctypes.Synchronized:
        effic_coef_i = effic_coef.value
    else:
        effic_coef_i = effic_coef
    if type(gener_loscoef_all)== sharedctypes.Synchronized:
        gener_loscoef_all_i = gener_loscoef_all.value
    else:
        gener_loscoef_all_i = gener_loscoef_all
    
    SK_norm_vol = look_SKvol(SK_norm_hgt)
    SK_dth_vol = look_SKvol(SK_dth_hgt)
    XK_norm_vol = look_XKvol(XK_norm_hgt)
    XK_dth_vol = look_XKvol(XK_dth_hgt)
    
    gross_head_avg = look_SKhgt(SK_dth_vol + (SK_norm_vol - SK_dth_vol)/2) - look_XKhgt(XK_dth_vol + (XK_norm_vol - XK_dth_vol)/2)
    gener_flow1 = installed_capcity_one_i * 10000 / effic_coef_i / gross_head_avg
    gener_flow2 = 0
    
    while abs(gener_flow1 - gener_flow2) > 0.001:
        head_loss = gener_loscoef_all_i * (gener_flow1 / installed_nums_i)**2 / 1000
        net_head = gross_head_avg - head_loss
        gener_flow2 = installed_capcity_i * 10000 / effic_coef_i / net_head
        gener_flow1 = gener_flow1 + 0.001
    
    head_loss = gener_loscoef_all_i * (gener_flow1 / installed_nums_i)**2 / 1000
    net_head = gross_head_avg - head_loss
    gener_flow2 = installed_capcity_i * 10000 / effic_coef_i / net_head
    gener_flow_avg = gener_flow2
    net_head_avg = net_head
    
    return(gross_head_avg, head_loss, gener_flow_avg, net_head_avg)
    
#### 蓄能量计算

def energy_storage_func(SK_dth_hgt,XK_dth_hgt):
    
    if type(SK_dth_hgt)== sharedctypes.Synchronized:
        SK_dth_hgt = SK_dth_hgt.value
    else:
        pass
    if type(XK_dth_hgt)== sharedctypes.Synchronized:
        XK_dth_hgt = XK_dth_hgt.value
    else:
        pass
    if type(effic_coef)== sharedctypes.Synchronized:
        effic_coef_i = effic_coef.value
    else:
        effic_coef_i = effic_coef
    if type(installed_nums)== sharedctypes.Synchronized:
        installed_nums_i = installed_nums.value
    else:
        installed_nums_i = installed_nums
    if type(installed_capcity)== sharedctypes.Synchronized:
        installed_capcity_i = installed_capcity.value
    else:
        installed_capcity_i = installed_capcity
    if type(gener_hours)== sharedctypes.Synchronized:
        gener_hours_i = gener_hours.value
    else:
        gener_hours_i = gener_hours
    
    calc_gener_secs = 0
    SK_dth_hgt0 = SK_dth_hgt
    XK_dth_hgt0 = XK_dth_hgt
    
    SK_dth_vol0 = look_SKvol(SK_dth_hgt0)
    adjust_vol_try = adjust_vol_try_func(SK_dth_hgt,XK_dth_hgt)
    SK_norm_hgt0 = round(look_SKhgt(look_SKvol(SK_dth_hgt0) + adjust_vol_try),1) ## 初始上库正常蓄水位
    SK_norm_hgt0 = SK_norm_hgt0 - 0.1
    
    actual_gener_secs = gener_hours_i * 3600 
       
    while calc_gener_secs < actual_gener_secs:
        SK_norm_hgt0 = SK_norm_hgt0 + 0.1
        
        calc_gener_secs = 0
        
        calc_slices_num = int(round(((SK_norm_hgt0 - SK_dth_hgt0) * 10),0))
        SK_heights = [None for i in range(calc_slices_num + 1)]
        SK_volumns = [None for i in range(calc_slices_num + 1)]
        SK_heights[0] = SK_norm_hgt0
        SK_volumns[0] = look_SKvol(SK_norm_hgt0)
        adjust_vol = SK_volumns[0] - SK_dth_vol0
        
        XK_heights = [None for i in range(calc_slices_num + 1)]
        XK_volumns = [None for i in range(calc_slices_num + 1)]
        XK_heights[0] = XK_dth_hgt0
        XK_volumns[0] = look_XKvol(XK_dth_hgt0)
        
        delta_volumns = [None for i in range(calc_slices_num + 1)]
        delta_volumns[0] = None
        
        gross_heads = [None for i in range(calc_slices_num + 1)]
        gross_heads[0] = None
        
        net_heads = [None for i in range(calc_slices_num + 1)]
        net_heads[0] = None
        
        head_losss = [None for i in range(calc_slices_num + 1)]
        head_losss[0] = None
        
        gener_flows = [None for i in range(calc_slices_num + 1)]
        gener_flows[0] = None
        
        gener_outputs = [None for i in range(calc_slices_num + 1)]
        gener_outputs[0] = None
        
        gener_secs = [None for i in range(calc_slices_num + 1)]
        gener_secs[0] = None
        
        gener_powers = [None for i in range(calc_slices_num + 1)]
        gener_powers[0] = None
        
        acumu_gener_powers = 0
        
        for i in range(calc_slices_num):
            SK_heights[i+1] = round(SK_heights[i]-0.1, 1)
            SK_volumns[i+1] = look_SKvol(SK_heights[i+1])
            delta_volumns[i+1] = SK_volumns[i] - SK_volumns[i+1]
            
            XK_volumns[i+1] = XK_volumns[i] + delta_volumns[i+1]
            XK_heights[i+1] = look_XKhgt(XK_volumns[i+1])
            
            gross_heads[i+1] = (SK_heights[i+1] + SK_heights[i])/2-(XK_heights[i+1] + XK_heights[i])/2
            net_heads[i+1] = loss_func(gross_heads[i+1])[0]
            head_losss[i+1] = loss_func(gross_heads[i+1])[1]
            gener_flows[i+1]  = loss_func(gross_heads[i+1])[2]

            gener_outputs[i+1] = effic_coef_i * gener_flows[i+1] * net_heads[i+1] / 10000
            gener_secs[i+1] = delta_volumns[i+1] * 10000 / gener_flows[i+1]
            gener_powers[i+1] = gener_outputs[i+1] * gener_secs[i+1] / 3600 
            
            calc_gener_secs = calc_gener_secs + gener_secs[i+1]
            acumu_gener_powers = acumu_gener_powers + gener_powers[i+1]
            
            result_list = [SK_heights, SK_volumns, delta_volumns, XK_heights, XK_volumns, gross_heads,
            head_losss, net_heads, gener_outputs, gener_flows, gener_secs, gener_powers]
            
        SK_params = [SK_norm_hgt0, SK_dth_hgt0, SK_volumns[0], SK_dth_vol0,adjust_vol] ##上库正常水位,上库死水位,上库正常库容,上库死库容,调节库容
        XK_params = [XK_heights[-1],XK_dth_hgt0,XK_volumns[-1],XK_volumns[0],adjust_vol]##下库正常水位,下库死水位,下库正常库容,下库死库容,调节库容
        gener_params = [installed_capcity_i,installed_nums_i,effic_coef_i,gener_hours_i,actual_gener_secs,calc_gener_secs,acumu_gener_powers] ## 装机容量，机组台数，出力系数，发电时间h，发电时间s,计算发电时间s,累计发电量
        
    return(SK_params, XK_params, gener_params, result_list)

### 抽水电能计算  
### 水头损失计算函数(公式法)
def pump_loss_func(gross_head):
    
    if type(gross_head)== sharedctypes.Synchronized:
        gross_head = gross_head.value
    else:
        pass
    if type(pump_effic_coef)== sharedctypes.Synchronized:
        pump_effic_coef_i = pump_effic_coef.value
    else:
        pump_effic_coef_i = pump_effic_coef
    if type(installed_nums)== sharedctypes.Synchronized:
        installed_nums_i = installed_nums.value
    else:
        installed_nums_i = installed_nums
    if type(pump_installed_capcity)== sharedctypes.Synchronized:
        pump_installed_capcity_i = pump_installed_capcity.value
    else:
        pump_installed_capcity_i = pump_installed_capcity
    if type(pump_loscoef_all)== sharedctypes.Synchronized:
        pump_loscoef_all_i = pump_loscoef_all.value
    else:
        pump_loscoef_all_i = pump_loscoef_all
    
    pump_flow1 = pump_installed_capcity_i * 10000 / pump_effic_coef_i / gross_head
    pump_flow2 = 0
    
    while abs(pump_flow1 - pump_flow2) > 0.001:
        head_loss = pump_loscoef_all_i * (pump_flow1 / installed_nums_i)**2 / 1000
        net_head = gross_head + head_loss
        pump_flow2 = pump_installed_capcity_i * 10000 / pump_effic_coef_i / net_head
        pump_flow1 = pump_flow1 - 0.001
        
    head_loss = pump_loscoef_all_i * (pump_flow1 / installed_nums_i)**2 / 1000
    net_head = gross_head + head_loss
    
    return(net_head,head_loss,pump_flow1)

### 计算最大扬程（全机组抽水）
def max_net_head_pump_func(SK_norm_hgt, XK_dth_hgt):
    
    if type(SK_norm_hgt)== sharedctypes.Synchronized:
        SK_norm_hgt = SK_norm_hgt.value
    else:
        pass
    if type(XK_dth_hgt)== sharedctypes.Synchronized:
        XK_dth_hgt = XK_dth_hgt.value
    else:
        pass
    if type(pump_installed_capcity)== sharedctypes.Synchronized:
        pump_installed_capcity_i = pump_installed_capcity.value
    else:
        pump_installed_capcity_i = pump_installed_capcity
    if type(pump_effic_coef)== sharedctypes.Synchronized:
        pump_effic_coef_i = pump_effic_coef.value
    else:
        pump_effic_coef_i = pump_effic_coef
    if type(installed_nums)== sharedctypes.Synchronized:
        installed_nums_i = installed_nums.value
    else:
        installed_nums_i = installed_nums
    if type(pump_loscoef_all)== sharedctypes.Synchronized:
        pump_loscoef_all_i = pump_loscoef_all.value
    else:
        pump_loscoef_all_i = pump_loscoef_all
    
    gross_head = SK_norm_hgt - XK_dth_hgt
    pump_flow1 = pump_installed_capcity_i * 10000 / pump_effic_coef_i / gross_head
    pump_flow2 = 0
    
    while abs(pump_flow1 - pump_flow2) > 0.001:
        head_loss = pump_loscoef_all_i * (pump_flow1 / installed_nums_i) **2 / 1000
        net_head = gross_head + head_loss
        pump_flow2 = pump_installed_capcity_i * 10000 / pump_effic_coef_i / net_head
        pump_flow1 = pump_flow1 - 0.001
    
    head_loss = pump_loscoef_all_i * (pump_flow1 / installed_nums_i) **2 / 1000
    net_head = gross_head + head_loss
    pump_flow2 = pump_installed_capcity_i * 10000 / pump_effic_coef_i / net_head
    pump_flow_max = pump_flow2
    net_head_max = net_head
    
    return(gross_head, head_loss, pump_flow_max, net_head_max)
    
### 计算最小扬程（一台机抽水）
def min_net_head_pump_func(SK_dth_hgt, XK_norm_hgt):
    
    if type(SK_dth_hgt)== sharedctypes.Synchronized:
        SK_dth_hgt = SK_dth_hgt.value
    else:
        pass
    if type(XK_norm_hgt)== sharedctypes.Synchronized:
        XK_norm_hgt = XK_norm_hgt.value
    else:
        pass
    if type(pump_effic_coef)== sharedctypes.Synchronized:
        pump_effic_coef_i = pump_effic_coef.value
    else:
        pump_effic_coef_i = pump_effic_coef
    if type(pump_installed_capcity_one)== sharedctypes.Synchronized:
        pump_installed_capcity_one_i = pump_installed_capcity_one.value
    else:
        pump_installed_capcity_one_i = pump_installed_capcity_one
    if type(pump_loscoef_one)== sharedctypes.Synchronized:
        pump_loscoef_one_i = pump_loscoef_one.value
    else:
        pump_loscoef_one_i = pump_loscoef_one
    
    if decim_plc == 1:
        gross_head = SK_dth_hgt - round(XK_norm_hgt+0.05,1)
    else:
        gross_head = SK_dth_hgt - round(XK_norm_hgt+0.5,0)
        
    pump_flow1 = pump_installed_capcity_one_i * 10000 / pump_effic_coef_i / gross_head
    pump_flow2 = 0
    
    while abs(pump_flow1 - pump_flow2) > 0.001:
        head_loss = pump_loscoef_one_i * pump_flow1 **2 / 1000
        net_head = gross_head + head_loss
        pump_flow2 = pump_installed_capcity_one_i * 10000 / pump_effic_coef_i / net_head
        pump_flow1 = pump_flow1 - 0.001
    
    head_loss = pump_loscoef_one_i * pump_flow1 **2 / 1000
    net_head = gross_head + head_loss
    pump_flow2 = pump_installed_capcity_one_i * 10000 / pump_effic_coef_i / net_head
    pump_flow_min = pump_flow2
    net_head_min = net_head
    
    return(gross_head, head_loss, pump_flow_min, net_head_min)

  
def pump_energy_func(SK_params,XK_params):
    
    if type(SK_params)== sharedctypes.SynchronizedArray:
        SK_params = list(SK_params)
    else:
        pass
    if type(XK_params)== sharedctypes.SynchronizedArray:
        XK_params = list(XK_params)
    else:
        pass
    if type(pump_effic_coef)== sharedctypes.Synchronized:
        pump_effic_coef_i = pump_effic_coef.value
    else:
        pump_effic_coef_i = pump_effic_coef
    if type(installed_nums)== sharedctypes.Synchronized:
        installed_nums_i = installed_nums.value
    else:
        installed_nums_i = installed_nums
    if type(pump_installed_capcity)== sharedctypes.Synchronized:
        pump_installed_capcity_i = pump_installed_capcity.value
    else:
        pump_installed_capcity_i = pump_installed_capcity
    if type(gener_hours)== sharedctypes.Synchronized:
        gener_hours_i = gener_hours.value
    else:
        gener_hours_i = gener_hours
    
    SK_norm_hgt = SK_params[0] 
    SK_norm_vol = SK_params[2]
    SK_dth_hgt = SK_params[1]
    SK_dth_vol = SK_params[3]
    XK_norm_hgt = XK_params[0]
    XK_norm_vol = XK_params[2]
         
    calc_pump_secs = 0
        
    calc_slices_num = int(round(((SK_norm_hgt - SK_dth_hgt) * 10),0))
    SK_heights = [None for i in range(calc_slices_num + 1)]
    SK_volumns = [None for i in range(calc_slices_num + 1)]
    SK_heights[0] = SK_dth_hgt
    SK_volumns[0] = SK_dth_vol
    adjust_vol = SK_norm_vol - SK_dth_vol
        
    XK_heights = [None for i in range(calc_slices_num + 1)]
    XK_volumns = [None for i in range(calc_slices_num + 1)]
    XK_heights[0] = XK_norm_hgt
    XK_volumns[0] = XK_norm_vol
        
    delta_volumns = [None for i in range(calc_slices_num + 1)]
    delta_volumns[0] = None
        
    gross_heads = [None for i in range(calc_slices_num + 1)]
    gross_heads[0] = None
        
    net_heads = [None for i in range(calc_slices_num + 1)]
    net_heads[0] = None
        
    head_losss = [None for i in range(calc_slices_num + 1)]
    head_losss[0] = None
        
    pump_flows = [None for i in range(calc_slices_num + 1)]
    pump_flows[0] = None
        
    pump_outputs = [None for i in range(calc_slices_num + 1)]
    pump_outputs[0] = None
        
    pump_secs = [None for i in range(calc_slices_num + 1)]
    pump_secs[0] = None
        
    pump_powers = [None for i in range(calc_slices_num + 1)]
    pump_powers[0] = None
        
    acumu_pump_powers = 0
    
    for i in range(calc_slices_num):
        SK_heights[i+1] = round(SK_heights[i]+0.1, 1)
        SK_volumns[i+1] = look_SKvol(SK_heights[i+1])
        delta_volumns[i+1] = SK_volumns[i+1] - SK_volumns[i]
            
        XK_volumns[i+1] = XK_volumns[i] - delta_volumns[i+1]
        XK_heights[i+1] = look_XKhgt(XK_volumns[i+1])
            
        gross_heads[i+1] = (SK_heights[i+1] + SK_heights[i])/2-(XK_heights[i+1] + XK_heights[i])/2
        net_heads[i+1] = pump_loss_func(gross_heads[i+1])[0]
        head_losss[i+1] = pump_loss_func(gross_heads[i+1])[1]
        pump_flows[i+1]  = pump_loss_func(gross_heads[i+1])[2]

        pump_outputs[i+1] = pump_effic_coef_i * pump_flows[i+1] * net_heads[i+1] /10000
        pump_secs[i+1] = delta_volumns[i+1] * 10000 / pump_flows[i+1]
        pump_powers[i+1] = pump_outputs[i+1] * pump_secs[i+1] / 3600
            
        calc_pump_secs = calc_pump_secs + pump_secs[i+1]
        acumu_pump_powers = acumu_pump_powers + pump_powers[i+1]
            
        result_list2 = [SK_heights, SK_volumns, delta_volumns, XK_heights, XK_volumns, gross_heads,
        head_losss, net_heads, pump_outputs, pump_flows, pump_secs, pump_powers]
            
    SK_params2 = [SK_heights[-1],SK_heights[0],SK_volumns[-1],SK_volumns[0],adjust_vol] ##上库正常水位,上库死水位,上库正常库容,上库死库容,调节库容
    XK_params2 = [XK_heights[0], XK_heights[-1],XK_volumns[0],XK_volumns[-1],adjust_vol]
    pump_params = [pump_installed_capcity_i, installed_nums_i,pump_effic_coef_i,gener_hours_i,gener_hours_i * 3600,
                   calc_pump_secs / 3600, acumu_pump_powers] ##装机容量,机组台数,入力系数,发电时间(h),发电时间(s),计算抽水时间(s),累计抽水电量
        
    return(SK_params2, XK_params2, pump_params, result_list2)
    
### 蓄能曲线计算
def energy_storage_func2(SK_dth_hgt,XK_dth_hgt,SK_norm_hgt):
    calc_gener_secs = 0
    
    calc_slices_num_i = int(round(((SK_norm_hgt - SK_dth_hgt) * 10),0))
    SK_heights = [None for i in range(calc_slices_num_i + 1)]
    SK_volumns = [None for i in range(calc_slices_num_i + 1)]
    SK_heights[0] = SK_norm_hgt
    SK_volumns[0] = look_SKvol(SK_norm_hgt)
        
    XK_heights = [None for i in range(calc_slices_num_i + 1)]
    XK_volumns = [None for i in range(calc_slices_num_i + 1)]
    XK_heights[0] = XK_dth_hgt
    XK_volumns[0] = look_XKvol(XK_dth_hgt)
        
    delta_volumns = [None for i in range(calc_slices_num_i + 1)]
    delta_volumns[0] = None
        
    gross_heads = [None for i in range(calc_slices_num_i + 1)]
    gross_heads[0] = None
        
    net_heads = [None for i in range(calc_slices_num_i + 1)]
    net_heads[0] = None
        
    head_losss = [None for i in range(calc_slices_num_i + 1)]
    head_losss[0] = None
        
    gener_flows = [None for i in range(calc_slices_num_i + 1)]
    gener_flows[0] = None
        
    gener_outputs = [None for i in range(calc_slices_num_i + 1)]
    gener_outputs[0] = None
        
    gener_secs = [None for i in range(calc_slices_num_i + 1)]
    gener_secs[0] = None
        
    gener_powers = [None for i in range(calc_slices_num_i + 1)]
    gener_powers[0] = None
        
    acumu_gener_powers = 0
        
    for i in range(calc_slices_num_i):
        SK_heights[i+1] = round(SK_heights[i]-0.1, 1)
        SK_volumns[i+1] = look_SKvol(SK_heights[i+1])
        delta_volumns[i+1] = SK_volumns[i] - SK_volumns[i+1]
        
        XK_volumns[i+1] = XK_volumns[i] + delta_volumns[i+1]
        XK_heights[i+1] = look_XKhgt(XK_volumns[i+1])
        
        gross_heads[i+1] = (SK_heights[i+1] + SK_heights[i])/2-(XK_heights[i+1] + XK_heights[i])/2
        net_heads[i+1] = loss_func(gross_heads[i+1])[0]
        head_losss[i+1] = loss_func(gross_heads[i+1])[1]
        gener_flows[i+1]  = loss_func(gross_heads[i+1])[2]
        
        gener_outputs[i+1] = effic_coef.value * gener_flows[i+1] * net_heads[i+1] / 10000
        gener_secs[i+1] = delta_volumns[i+1] * 10000 / gener_flows[i+1]
        gener_powers[i+1] = gener_outputs[i+1] * gener_secs[i+1] / 3600 
            
        calc_gener_secs = calc_gener_secs + gener_secs[i+1]
        acumu_gener_powers = acumu_gener_powers + gener_powers[i+1]
        
    return(calc_gener_secs/3600, acumu_gener_powers,XK_heights[-1],XK_volumns[-1])

def energy_storage_curve(SK_dth_hgt,XK_dth_hgt,SK_norm_hgt):
    
    calc_slices_num = int(round(((SK_norm_hgt - SK_dth_hgt) * 10),0))
    
    SK_heights = [None for i in range(calc_slices_num + 1)]
    SK_volumns = [None for i in range(calc_slices_num + 1)]
    SK_heights[0] = SK_dth_hgt
    SK_volumns[0] = look_SKvol(SK_dth_hgt)
        
    XK_heights = [None for i in range(calc_slices_num + 1)]
    XK_volumns = [None for i in range(calc_slices_num + 1)]
    XK_heights[0] = XK_dth_hgt
    XK_volumns[0] = look_XKvol(XK_dth_hgt)
    
    gener_outputs = [installed_capcity.value for i in range(calc_slices_num + 1)]
        
    gener_secs = [None for i in range(calc_slices_num + 1)]
    gener_secs[0] = 0
        
    gener_powers = [None for i in range(calc_slices_num + 1)]
    gener_powers[0] = 0
    
    for i in range(calc_slices_num):
        SK_heights[i+1] = round(SK_heights[i] + 0.1,1)
        SK_volumns[i+1] = look_SKvol(SK_heights[i+1])
        
        gener_secs[i+1],gener_powers[i+1],XK_heights[i+1],XK_volumns[i+1] = energy_storage_func2(SK_dth_hgt,
                                                                                                XK_dth_hgt,
                                                                                                SK_heights[i+1])
        print('正在计算：水位' + str(SK_heights[i+1]) + '对应蓄能量！')
        
    result_list = [SK_heights,SK_volumns,XK_heights,XK_volumns,gener_outputs,gener_secs,gener_powers]
    return (result_list)
    
if __name__ == '__main__':
    
    ### TK window
    window = tk.Tk()
    window.title('抽水蓄能调算程序')
    window.maxsize(1044,720)
    window.minsize(1044,720)
    
    canvas = tk.Canvas(window, height=720, width=1024)
    canvas.pack(side='top')
    cover_img = tk.PhotoImage(file = 'cover2.gif')
    cover_image = canvas.create_image(5,10,anchor='nw',image = cover_img)
    
    #### TK输入文件路径
    krqx_filepath = tk.StringVar()
    krqx_filepath.set('请输入库容曲线文件(excel格式):')
    krqx_file_entry = tk.Entry(window, width = 80, textvariable = krqx_filepath)
    krqx_file_entry.place(x=130,y=450)
    def choose_file():
        selectFileName = tk.filedialog.askopenfilename(title='选择文件') 
        krqx_filepath.set(selectFileName)
        window.destroy()
    sumbit_btn = tk.Button(window,text="选择文件",bg='lightgray',command = choose_file)
    sumbit_btn.place(x=720,y=445)
    
    label_cr = tk.Label(window,text='Copyright  ©  huangzq681@gmail.com',fg="gray",
                      font = ('Arial',10, "italic"))
    label_cr.place(x=425,y=700)
       
    ##### TK上下库参数
    label_updw = tk.Label(window, text='上下库参数',font=('微软雅黑',14))
    label_updw.place(x=65,y=485)
    
    label_updth = tk.Label(window, text='上库死水位：')
    label_updth.place(x=75,y=515)
    label_dwdth = tk.Label(window, text='下库死水位：')
    label_dwdth.place(x=75,y=545)
    up_dth_hgt = tk.StringVar()
    up_dth_hgt_e = tk.Entry(window, width = 10, textvariable = up_dth_hgt)
    up_dth_hgt_e.place(x=160,y=515)
    dw_dth_hgt = tk.StringVar()
    dw_dth_hgt_e = tk.Entry(window, width = 10, textvariable = dw_dth_hgt)
    dw_dth_hgt_e.place(x=160,y=545)
    canvas.create_rectangle(55, 500, 265, 570, dash=(4, 4),fill = '')
    
    ##### TK机组参数
    label_mech = tk.Label(window, text='机组参数',font=('微软雅黑',16))
    label_mech.place(x=285,y=485)
    
    lb_inst_cap = tk.Label(window, text='装机容量(万kw)：')
    lb_inst_cap.place(x=295,y=515)
    lb_inst_num = tk.Label(window, text='机组台数(台)：')
    lb_inst_num.place(x=295,y=545)
    lb_gener_coef = tk.Label(window, text='出力系数：')
    lb_gener_coef.place(x=295,y=575)
    lb_gener_hour = tk.Label(window, text='发电时数(h)：')
    lb_gener_hour.place(x=295,y=605)
    inst_cap = tk.StringVar()
    inst_cap_e = tk.Entry(window, width = 10, textvariable = inst_cap)
    inst_cap_e.place(x=400,y=515)
    inst_num = tk.StringVar()
    inst_num_e = tk.Entry(window, width = 10, textvariable = inst_num)
    inst_num_e.place(x=400,y=545)
    gener_coef = tk.StringVar()
    gener_coef.set('')
    gener_coef_e = tk.Entry(window, width = 10, textvariable = gener_coef)
    gener_coef_e.place(x=400,y=575)
    gener_hour = tk.StringVar()
    gener_hour_e = tk.Entry(window, width = 10, textvariable = gener_hour)
    gener_hour_e.place(x=400,y=605)
    canvas.create_rectangle(275, 500, 505, 640, dash=(4, 4),fill = '')
    
    ##### TK水头损失参数
    label_loss = tk.Label(window, text='水头损失系数',font=('微软雅黑',14))
    label_loss.place(x=515,y=485)
    
    lb_generloss_all = tk.Label(window, text='发电工况(全机)：')
    lb_generloss_all.place(x=535,y=515)
    lb_generloss_one = tk.Label(window, text='发电工况(单机)：')
    lb_generloss_one.place(x=535,y=545)
    lb_pumploss_all = tk.Label(window, text='抽水工况(全机)：')
    lb_pumploss_all.place(x=535,y=575)
    lb_pumploss_one = tk.Label(window, text='抽水工况(单机)：')
    lb_pumploss_one.place(x=535,y=605)
    generloss_all = tk.StringVar()
    generloss_all.set('')
    generloss_all_e = tk.Entry(window, width = 14, textvariable = generloss_all)
    generloss_all_e.place(x=640,y=515)
    generloss_one = tk.StringVar()
    generloss_one.set('')
    generloss_one_e = tk.Entry(window, width = 14, textvariable = generloss_one)
    generloss_one_e.place(x=640,y=545)
    pumploss_all = tk.StringVar()
    pumploss_all.set('')
    pumploss_all_e = tk.Entry(window, width = 14, textvariable = pumploss_all)
    pumploss_all_e.place(x=640,y=575)
    pumploss_one = tk.StringVar()
    pumploss_one.set('')
    pumploss_one_e = tk.Entry(window, width = 14, textvariable = pumploss_one)
    pumploss_one_e.place(x=640,y=605)
    canvas.create_rectangle(515, 500, 745, 640, dash=(4, 4),fill = '')
    
    
    ##### TK保留小数位
    label_res = tk.Label(window, text='计算小数位',font=('微软雅黑',14))
    label_res.place(x=65,y=575)
    lb_generloss_all = tk.Label(window, text='保留小数位：')
    lb_generloss_all.place(x=75,y=605)
    cal_res = tk.IntVar()
    cal_res.set(1)
    cal_res_value1 = tk.Radiobutton(window, text="0.0", value=1, variable=cal_res,font = ('Arial'))
    cal_res_value2 = tk.Radiobutton(window, text="0", value=2, variable=cal_res,font = ('Arial'))
    cal_res_value1.place(x=160,y=605)
    cal_res_value2.place(x=220,y=605)
    canvas.create_rectangle(55, 590, 265, 640, dash=(4, 4),fill = '')
    
     ##### TK是否计算蓄能曲线
    label_xnqx = tk.Label(window, text='蓄能曲线计算',font=('微软雅黑',14))
    label_xnqx.place(x=765,y=485)
    label_cal_xnqx = tk.Label(window, text='是否计算蓄能曲线？')
    label_cal_xnqx.place(x=775,y=515)
    cal_xnqx = tk.IntVar()
    cal_xnqx.set(2)
    cal_xnqx_value1 = tk.Radiobutton(window, text="是", value=1, variable=cal_xnqx,font = ('Arial'))
    cal_xnqx_value2 = tk.Radiobutton(window, text="否", value=2, variable=cal_xnqx,font = ('Arial'))
    cal_xnqx_value1.place(x=895,y=515)
    cal_xnqx_value2.place(x=940,y=515)
    label_xnqx1 = tk.Label(window,text='(默认不计算蓄能曲线，计算耗时较长)',fg="gray",
                      font = ('Arial',8, "italic"))
    label_xnqx1.place(x=775,y=545)
    
    window.mainloop()
    
    #### TK运行
    
    #### 上下库库容曲线
    krqx_file = krqx_filepath.get()
    try:
        krqx = xlrd.open_workbook(krqx_file)
    except IOError:
        print ("Error: 文件读取失败")
    else:
        print ("文件读取成功！")
    
    SK_krqx = krqx.sheet_by_index(0)
    XK_krqx = krqx.sheet_by_index(1)
    
    SK_num = SK_krqx.nrows-1
    XK_num = XK_krqx.nrows-1
    SK_hgts = [SK_krqx.row_values(i+1)[0] for i in range(SK_num)]
    SK_vols = [SK_krqx.row_values(i+1)[1] for i in range(SK_num)]
    XK_hgts = [XK_krqx.row_values(i+1)[0] for i in range(XK_num)]
    XK_vols = [XK_krqx.row_values(i+1)[1] for i in range(XK_num)]
    
    SK_hgts = Array('f',SK_hgts)
    XK_hgts = Array('f',XK_hgts)
    SK_vols = Array('f',SK_vols)
    XK_vols = Array('f',XK_vols)
    
    loss_params = krqx.sheet_by_index(2)
    station_name = loss_params.row_values(0)[1]
    pump_effic_coef_input = loss_params.row_values(2)[1]
    gener_coef_input = loss_params.row_values(1)[1]
    
    ### TK window
    window = tk.Tk()
    window.title('抽水蓄能调算程序')
    window.maxsize(1044,720)
    window.minsize(1044,720)
    
    canvas = tk.Canvas(window, height=720, width=1024)
    canvas.pack(side='top')
    cover_img = tk.PhotoImage(file = 'cover2.gif')
    cover_image = canvas.create_image(5,10,anchor='nw',image = cover_img)
    
    #### TK输入文件路径
    krqx_filepath2 = tk.StringVar()
    krqx_filepath2.set(krqx_file)
    krqx_file_entry = tk.Entry(window, width = 80, textvariable = krqx_filepath2)
    krqx_file_entry.place(x=130,y=450)
    sumbit_btn = tk.Button(window,text="选择文件",bg='lightgray')
    sumbit_btn.place(x=720,y=445)
    
    label_cr = tk.Label(window,text='Copyright  ©  huangzq681@gmail.com',fg="gray",
                      font = ('Arial',10, "italic"))
    label_cr.place(x=425,y=700)
       
    ##### TK上下库参数
    label_updw = tk.Label(window, text='上下库参数',font=('微软雅黑',14))
    label_updw.place(x=65,y=485)
    
    label_updth = tk.Label(window, text='上库死水位：')
    label_updth.place(x=75,y=515)
    label_dwdth = tk.Label(window, text='下库死水位：')
    label_dwdth.place(x=75,y=545)
    up_dth_hgt = tk.StringVar()
    up_dth_hgt_e = tk.Entry(window, width = 10, textvariable = up_dth_hgt)
    up_dth_hgt_e.place(x=160,y=515)
    dw_dth_hgt = tk.StringVar()
    dw_dth_hgt_e = tk.Entry(window, width = 10, textvariable = dw_dth_hgt)
    dw_dth_hgt_e.place(x=160,y=545)
    canvas.create_rectangle(55, 500, 265, 570, dash=(4, 4),fill = '')
    
    ##### TK机组参数
    label_mech = tk.Label(window, text='机组参数',font=('微软雅黑',16))
    label_mech.place(x=285,y=485)
    
    lb_inst_cap = tk.Label(window, text='装机容量(万kw)：')
    lb_inst_cap.place(x=295,y=515)
    lb_inst_num = tk.Label(window, text='机组台数(台)：')
    lb_inst_num.place(x=295,y=545)
    lb_gener_coef = tk.Label(window, text='出力系数：')
    lb_gener_coef.place(x=295,y=575)
    lb_gener_hour = tk.Label(window, text='发电时数(h)：')
    lb_gener_hour.place(x=295,y=605)
    inst_cap = tk.StringVar()
    inst_cap_e = tk.Entry(window, width = 10, textvariable = inst_cap)
    inst_cap_e.place(x=400,y=515)
    inst_num = tk.StringVar()
    inst_num_e = tk.Entry(window, width = 10, textvariable = inst_num)
    inst_num_e.place(x=400,y=545)
    gener_coef = tk.StringVar()
    gener_coef.set(str(gener_coef_input))
    gener_coef_e = tk.Entry(window, width = 10, textvariable = gener_coef)
    gener_coef_e.place(x=400,y=575)
    gener_hour = tk.StringVar()
    gener_hour_e = tk.Entry(window, width = 10, textvariable = gener_hour)
    gener_hour_e.place(x=400,y=605)
    canvas.create_rectangle(275, 500, 505, 640, dash=(4, 4),fill = '')
    
    ##### TK水头损失参数
    label_loss = tk.Label(window, text='水头损失系数',font=('微软雅黑',14))
    label_loss.place(x=515,y=485)
    
    lb_generloss_all = tk.Label(window, text='发电工况(全机)：')
    lb_generloss_all.place(x=535,y=515)
    lb_generloss_one = tk.Label(window, text='发电工况(单机)：')
    lb_generloss_one.place(x=535,y=545)
    lb_pumploss_all = tk.Label(window, text='抽水工况(全机)：')
    lb_pumploss_all.place(x=535,y=575)
    lb_pumploss_one = tk.Label(window, text='抽水工况(单机)：')
    lb_pumploss_one.place(x=535,y=605)
    generloss_all = tk.StringVar()
    generloss_all.set('')
    generloss_all_e = tk.Entry(window, width = 14, textvariable = generloss_all)
    generloss_all_e.place(x=640,y=515)
    generloss_one = tk.StringVar()
    generloss_one.set('')
    generloss_one_e = tk.Entry(window, width = 14, textvariable = generloss_one)
    generloss_one_e.place(x=640,y=545)
    pumploss_all = tk.StringVar()
    pumploss_all.set('')
    pumploss_all_e = tk.Entry(window, width = 14, textvariable = pumploss_all)
    pumploss_all_e.place(x=640,y=575)
    pumploss_one = tk.StringVar()
    pumploss_one.set('')
    pumploss_one_e = tk.Entry(window, width = 14, textvariable = pumploss_one)
    pumploss_one_e.place(x=640,y=605)
    canvas.create_rectangle(515, 500, 745, 640, dash=(4, 4),fill = '')
    
    generloss_all.set(str(loss_params.row_values(5)[1]))
    generloss_one.set(str(loss_params.row_values(6)[1]))
    pumploss_all.set(str(loss_params.row_values(7)[1]))
    pumploss_one.set(str(loss_params.row_values(8)[1]))  
    
    generloss_all_e.config(textvariable = generloss_all)
    generloss_one_e.config(textvariable = generloss_one)
    pumploss_all_e.config(textvariable = pumploss_all)
    pumploss_one_e.config(textvariable = pumploss_one)
    
    
    ##### TK保留小数位
    label_res = tk.Label(window, text='计算小数位',font=('微软雅黑',14))
    label_res.place(x=65,y=575)
    lb_generloss_all = tk.Label(window, text='保留小数位：')
    lb_generloss_all.place(x=75,y=605)
    cal_res = tk.IntVar()
    cal_res.set(1)
    cal_res_value1 = tk.Radiobutton(window, text="0.0", value=1, variable=cal_res,font = ('Arial'))
    cal_res_value2 = tk.Radiobutton(window, text="0", value=2, variable=cal_res,font = ('Arial'))
    cal_res_value1.place(x=160,y=605)
    cal_res_value2.place(x=220,y=605)
    canvas.create_rectangle(55, 590, 265, 640, dash=(4, 4),fill = '')
    
     ##### TK是否计算蓄能曲线
    label_xnqx = tk.Label(window, text='蓄能曲线计算',font=('微软雅黑',14))
    label_xnqx.place(x=765,y=485)
    label_cal_xnqx = tk.Label(window, text='是否计算蓄能曲线？')
    label_cal_xnqx.place(x=775,y=515)
    cal_xnqx = tk.IntVar()
    cal_xnqx.set(2)
    cal_xnqx_value1 = tk.Radiobutton(window, text="是", value=1, variable=cal_xnqx,font = ('Arial'))
    cal_xnqx_value2 = tk.Radiobutton(window, text="否", value=2, variable=cal_xnqx,font = ('Arial'))
    cal_xnqx_value1.place(x=895,y=515)
    cal_xnqx_value2.place(x=940,y=515)
    label_xnqx1 = tk.Label(window,text='(默认不计算蓄能曲线，计算耗时较长)',fg="gray",
                      font = ('Arial',8, "italic"))
    label_xnqx1.place(x=775,y=545)
    
        
    #### TK输出文件目录
    res_filepath = tk.StringVar()
    res_filepath.set('输出文件路径:')
    res_file_entry = tk.Entry(window, width = 80, textvariable = res_filepath)
    res_file_entry.place(x=130,y=650)
    
    def res_path():
        selectFileName = tk.filedialog.askdirectory(title='选择文件路径') 
        res_filepath.set(selectFileName)
    
    start_img = tk.PhotoImage(file = 'start2.png')
    
    res_btn = tk.Button(window,text="输出文件目录",bg='lightgray',command = res_path)
    res_btn.place(x=720,y=645)
    start_btn = tk.Button(master = window,image=start_img,command = window.destroy)
    start_btn.place(x=950,y=650)
    
    window.mainloop()
    
    ##### 相关参数
    SK_dth_height_input = Value('f',float(up_dth_hgt.get()))
    XK_dth_height_input = Value('f',float(dw_dth_hgt.get()))
    decim_plc = Value('i',cal_res.get())
    gener_loscoef_all = Value('f',float(generloss_all.get()))
    gener_loscoef_one = Value('f',float(generloss_one.get()))
    installed_capcity = Value('i',int(inst_cap.get())) # 万kw
    installed_nums = Value('i',int(inst_num.get()))
    installed_capcity_one = Value('f',installed_capcity.value / installed_nums.value)
    effic_coef = Value('f',float(gener_coef.get()))
    gener_hours = Value('f',float(gener_hour.get())) # h
    min_grossloss_div_netloss = Value('f',1.02)
    max_grossloss_div_netloss = Value('f',1.04)
    energy_storage = Value('f',installed_capcity.value * gener_hours.value)
    
    pump_loscoef_all = Value('f',float(pumploss_all.get()))
    pump_loscoef_one = Value('f',float(pumploss_one.get()))
    pump_installed_capcity = Value('f',installed_capcity.value * 1.083)
    pump_installed_capcity_one = Value('f',pump_installed_capcity.value / installed_nums.value)
    pump_effic_coef = Value('f',pump_effic_coef_input)
    
    ########## 蓄能量计算
    time1_start = time.time()
    SK_params, XK_params, gener_params, result_list = energy_storage_func(SK_dth_height_input, XK_dth_height_input)
    SK_params = Array('f',SK_params)
    XK_params = Array('f',XK_params)
    gener_params = Array('f',gener_params)
    net_head_max = max_net_head(SK_params[0],XK_params[1]) ##上库正常水位,上库死水位,上库正常库容,上库死库容,调节库容
    net_head_min = min_net_head(XK_params[0],SK_params[1])
    net_head_avg = average_net_head(SK_params[0],SK_params[1],
                                    XK_params[0],XK_params[1])
    head_loss_list = {"全机发电损失系数":gener_loscoef_all, "单机发电损失系数":gener_loscoef_one,
                      "最大毛水头":net_head_max[0], "最大净水头对应水损":net_head_max[1],
                       "最大净水头对应发电流量":net_head_max[2], "最大净水头":net_head_max[3],
                       "最小毛水头":net_head_min[0], "最小净水头对应水损":net_head_min[1],
                       "最小净水头对应发电流量":net_head_min[2], "最小净水头":net_head_min[3],
                       "平均毛水头":net_head_avg[0], "平均净水头对应水损":net_head_avg[1],
                       "平均净水头对应发电流量":net_head_avg[2], "平均净水头":net_head_avg[3]}
    
    time1_end = time.time()
    print('蓄能量计算时间:'+str(time1_end - time1_start))
    
    ########## 抽水电能计算
    time2_start = time.time()
    SK_params2, XK_params2, pump_params, result_list2 = pump_energy_func(SK_params,XK_params)
    max_net_head_pump = max_net_head_pump_func(SK_params[0],XK_params[1])
    min_net_head_pump = min_net_head_pump_func(SK_params[1],XK_params[0])
    
    head_loss_list_pump = {"全机抽水损失系数":pump_loscoef_all, "单机抽水损失系数":pump_loscoef_one,
                           "最大毛水头":max_net_head_pump[0], "最大扬程对应水损":max_net_head_pump[1],
                           "最大扬程对应发电流量":max_net_head_pump[2], "最大扬程":max_net_head_pump[3],
                           "最小毛水头":min_net_head_pump[0], "最小扬程对应水损":min_net_head_pump[1],
                           "最小扬程对应发电流量":min_net_head_pump[2], "最小扬程":min_net_head_pump[3]}   
    
    time2_end = time.time()
    print('抽水电能计算时间:'+str(time2_end - time2_start))
    
    ########### 蓄能曲线计算
    if cal_xnqx.get() == 1:
        energy_storage_curve1 = energy_storage_curve(SK_params[1],XK_params[1],SK_params[0])  
    else:
        pass

    ### 写文档
    print("正在写文档！")
    
    work_dir = res_filepath.get()
#    to_file_name = r"蓄能量调节计算.xlsx"
    to_file_name = r".xlsx"
    to_file_path = work_dir + "/" + station_name + '-' + str(installed_capcity.value)+\
    '万kW-' + str(gener_hours.value) + 'h-上'+ str(round(SK_params[0],1)) + '(' + str(SK_params[1]) + ')-下' + str(round(XK_params[0],1)) +\
    '(' + str(XK_params[1]) + ')-' + '水头变幅' + str(round(head_loss_list_pump['最大扬程']/head_loss_list['最小净水头'],3)) + to_file_name
    
    book = xlsxwriter.Workbook(to_file_path, options={  # 全局设置
            'strings_to_numbers': True,  # str 类型数字转换为 int 数字
            'strings_to_urls': False,  # 自动识别超链接
            'constant_memory': False,  # 连续内存模式 (True 适用于大数据量输出)
            'default_format_properties': {
                'font_name': '微软雅黑',  # 字体. 默认值 "Arial"
                'font_size': 10,  # 字号. 默认值 11
                # 'bold': False,  # 字体加粗
                # 'border': 1,  # 单元格边框宽度. 默认值 0
                'align': 'center',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                # 'text_wrap': False,  # 单元格内是否自动换行
                # ...
            },
        })
    
    sheet1 = book.add_worksheet(u"成果汇总")
    sheet2 = book.add_worksheet(u"蓄能量计算表")
    sheet3 = book.add_worksheet(u"抽水电能计算表")
    sheet4 = book.add_worksheet(u"蓄能量曲线")
    sheet5 = book.add_worksheet(u"上库库容曲线")
    sheet6 = book.add_worksheet(u"下库库容曲线")
    
    # 成果汇总
    fmt0 = book.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    sheet1.merge_range(0, 0, 0, 1, u"装机容量", cell_format = fmt0)
    sheet1.merge_range(0, 2, 0, 3, str(installed_capcity.value)+'万kW', cell_format = fmt0)
    sheet1.merge_range(1, 0, 1, 1, u"装机台数", cell_format = fmt0)
    sheet1.merge_range(1, 2, 1, 3, installed_nums.value, cell_format = fmt0)
    
    fmt1 = book.add_format({'font_size': 12, 'bold': True, 'border': 1, 'fg_color': '#FFDAB9'})
    sheet1.write(3, 0, u"项目", fmt1)
    sheet1.write(3, 1, u"单位", fmt1)
    sheet1.write(3, 2, u"上库", fmt1)
    sheet1.write(3, 3, u"下库", fmt1)
    
    fmt2 = book.add_format({'font_size': 10, 'bold': True, 'border': 1, 'fg_color': '#E6E6FA'})
    xiangmu = [u'正常蓄水位',u'相应库容',u'死水位',u'死库容',u'调节库容',u'消落深度',
               u'最大毛水头',u'最小毛水头',u'最大净水头',u'最小净水头',u'最大扬程',
               u'最小扬程',u'最大扬程/最小净水头',u'平均毛水头',u'平均净水头',u'满发小时数']
    for i in range(len(xiangmu)):
        sheet1.write(4+i, 0, xiangmu[i], fmt2)
        
    fmt3 = book.add_format({'font_name':'Times New Roman','italic': True, 'border': 1})
    danwei = ['m','万m3','m','万m3','万m3','m','m','m','m','m','m','m','','m','m','h']
    for i in range(len(danwei)):
        sheet1.write(4+i, 1, danwei[i], fmt3)
        
    SK_params3 = [SK_params[0], SK_params[2], 
                  SK_params[1], SK_params[3], 
                  SK_params[4], SK_params[0]-SK_params[1]]
    
    if decim_plc.value == 1:
        XK_norm_hgt_app = round(XK_params[0]+0.05,1)
    else:
        XK_norm_hgt_app = round(XK_params[0]+0.5,0)
    XK_norm_vol_app = look_XKvol(XK_norm_hgt_app)
      
    XK_params3 = [XK_norm_hgt_app, XK_norm_vol_app, 
                  XK_params[1], XK_params[3], 
                  XK_params[4],XK_norm_hgt_app-XK_params[1]]
    
    fmt4 = book.add_format({'font_name':'Times New Roman','num_format':'0.00', 'border': 1})
    fmt5 = book.add_format({'font_name':'Times New Roman','num_format':'0.0', 'border': 1})
    fmt6 = book.add_format({'font_name':'Times New Roman','num_format':'0.000', 'border': 1, 'bold': True})
    fmt8 = book.add_format({'font_name':'Times New Roman','num_format':'0', 'border': 1})
    
    sheet1.write(4, 2, SK_params3[0], fmt4)
    sheet1.write(6, 2, SK_params3[2], fmt4)
    sheet1.write(9, 2, SK_params3[5], fmt4)
    sheet1.write(5, 2, SK_params3[1], fmt5)
    sheet1.write(7, 2, SK_params3[3], fmt5)
    sheet1.write(8, 2, SK_params3[4], fmt5)
    
    sheet1.write(4, 3, XK_params3[0], fmt4)
    sheet1.write(6, 3, XK_params3[2], fmt4)
    sheet1.write(9, 3, XK_params3[5], fmt4)
    sheet1.write(5, 3, XK_params3[1], fmt5)
    sheet1.write(7, 3, XK_params3[3], fmt5)
    sheet1.write(8, 3, XK_params3[1]-XK_params3[3], fmt5)
    
    sheet1.merge_range(10,2,10,3,head_loss_list['最大毛水头'],fmt4)
    sheet1.merge_range(11,2,11,3,head_loss_list['最小毛水头'],fmt4)
    sheet1.merge_range(12,2,12,3,head_loss_list['最大净水头'],fmt4)
    sheet1.merge_range(13,2,13,3,head_loss_list['最小净水头'],fmt4) 
    sheet1.merge_range(14,2,14,3,head_loss_list_pump['最大扬程'],fmt4)   
    sheet1.merge_range(15,2,15,3,head_loss_list_pump['最小扬程'],fmt4) 
    sheet1.merge_range(16,2,16,3,head_loss_list_pump['最大扬程']/head_loss_list['最小净水头'],fmt6) 
    sheet1.merge_range(17,2,17,3,head_loss_list['平均毛水头'],fmt4)
    sheet1.merge_range(18,2,18,3,head_loss_list['平均净水头'],fmt4)
    sheet1.merge_range(19,2,19,3,gener_hours.value,fmt4)
    
    sheet1.set_column(0,0,16)
    
    fmt7 = book.add_format({'font_name':'Arial','italic': True, 'fg_color': '#E8E8E8', 'font_color': '#778899'})
    sheet1.merge_range(20,0,20,3,"Copyright © huangzq681@gmail.com", fmt7)
    
    # 蓄能量计算表
    sheet2.merge_range(0, 0, 0, 2, u"上库特征值", cell_format = fmt1)
    sheet2.merge_range(0, 3, 0, 5, u"下库特征值", cell_format = fmt1)
    sheet2.merge_range(0, 6, 0, 8, u"发电参数", cell_format = fmt1)
    sheet2.merge_range(0, 9, 0, 12, u"水头损失", cell_format = fmt1)
    
    SK_params_names = ['正常蓄水位 (m)','死水位 (m)','正常库容 (万m3)','死库容 (万m3)','调节库容 (万m3)',
                       '正常蓄水位 (采用,m)','正常库容 (采用,万m3)','调节库容 (采用,万m3)']
    for i in range(len(SK_params_names)):
        sheet2.merge_range(1+i, 0, 1+i, 1, SK_params_names[i], fmt2)
        
    sheet2.write(1,2,SK_params[0],fmt5)
    sheet2.write(2,2,SK_params[1],fmt5)
    sheet2.write(3,2,SK_params[2],fmt4)
    sheet2.write(4,2,SK_params[3],fmt4)
    sheet2.write(5,2,SK_params[4],fmt4)
    sheet2.write(6,2,SK_params[0],fmt5)
    sheet2.write(7,2,SK_params[2],fmt4)
    sheet2.write(8,2,SK_params[4],fmt4)
    
    XK_params_names = ['正常蓄水位 (m)','死水位 (m)','正常库容 (万m3)','死库容 (万m3)','调节库容 (万m3)',
                       '正常蓄水位 (采用,m)','正常库容 (采用,万m3)','调节库容 (采用,万m3)']
    for i in range(len(XK_params_names)):
        sheet2.merge_range(1+i, 3, 1+i, 4, XK_params_names[i], fmt2)
        
    sheet2.write(1,5,XK_params[0],fmt5)
    sheet2.write(2,5,XK_params[1],fmt5)
    sheet2.write(3,5,XK_params[2],fmt4)
    sheet2.write(4,5,XK_params[3],fmt4)
    sheet2.write(5,5,XK_params[4],fmt4)
    sheet2.write(6,5,XK_norm_hgt_app,fmt5)
    sheet2.write(7,5,XK_norm_vol_app,fmt4)
    sheet2.write(8,5,XK_norm_vol_app - XK_params[3],fmt4)
    
    gener_params_names = ['装机容量 (万kW)','出力系数','发电时间 (h)','发电时间 (s)',
                          '计算时间 (s)','发电量 (万kW.h)','机组台数 (台)','单机容量 (万kW)']
    for i in range(len(gener_params_names)):
        sheet2.merge_range(1+i, 6, 1+i, 7, gener_params_names[i], fmt2)
    
    sheet2.write(1,8,gener_params[0],fmt8)
    sheet2.write(2,8,gener_params[2],fmt4)
    sheet2.write(3,8,gener_params[3],fmt4)
    sheet2.write(4,8,gener_params[4],fmt8)
    sheet2.write(5,8,gener_params[5],fmt8)
    sheet2.write(6,8,gener_params[6],fmt8)
    sheet2.write(7,8,gener_params[1],fmt8)
    sheet2.write(8,8,gener_params[0]/gener_params[1],fmt5)
    
    sheet2.set_column(9,12,15)
    sheet2.write(1,9,"全机发电损失系数",fmt2)
    sheet2.write(1,10,"单机发电损失系数",fmt2)
    sheet2.write(1,11,"最大水头比 (毛/净)",fmt2)
    sheet2.write(1,12,"最小水头比 (毛/净)",fmt2)
    sheet2.write(2,9,gener_loscoef_all.value,fmt6)
    sheet2.write(2,10,gener_loscoef_one.value,fmt6)
    sheet2.write(2,11,max_grossloss_div_netloss.value,fmt4)
    sheet2.write(2,12,min_grossloss_div_netloss.value,fmt4)
    
    head_loss_name1 = [u'最大毛水头 (m)',u'发电流量 (m3/s)',u'水头损失 (m)',u'最大净水头 (m)']
    head_loss_name2 = [u'最小毛水头 (m)',u'发电流量 (m3/s)',u'水头损失 (m)',u'最小净水头 (m)']
    head_loss_name3 = [u'平均毛水头 (m)',u'发电流量 (m3/s)',u'水头损失 (m)',u'平均净水头 (m)']
    head_loss_list1 = [head_loss_list['最大毛水头'],head_loss_list['最大净水头对应发电流量'],
                       head_loss_list['最大净水头对应水损'],head_loss_list['最大净水头']]
    head_loss_list2 = [head_loss_list['最小毛水头'],head_loss_list['最小净水头对应发电流量'],
                       head_loss_list['最小净水头对应水损'],head_loss_list['最小净水头']]
    head_loss_list3 = [head_loss_list['平均毛水头'],head_loss_list['平均净水头对应发电流量'],
                       head_loss_list['平均净水头对应水损'],head_loss_list['平均净水头']]
    
    fmt9 = book.add_format({'font_size': 10, 'bold': True, 'border': 1, 'fg_color': '#FAEBD7'})
    
    for i in range(len(head_loss_name1)):
        sheet2.write(3, 9+i, head_loss_name1[i], fmt9)
    for i in range(len(head_loss_name2)):
        sheet2.write(5, 9+i, head_loss_name2[i], fmt9)
    for i in range(len(head_loss_name3)):
        sheet2.write(7, 9+i, head_loss_name3[i], fmt9)
    for i in range(len(head_loss_list1)):
        sheet2.write(4, 9+i, head_loss_list1[i], fmt4)
    for i in range(len(head_loss_list2)):
        sheet2.write(6, 9+i, head_loss_list2[i], fmt4)
    for i in range(len(head_loss_list3)):
        sheet2.write(8, 9+i, head_loss_list3[i], fmt4)
        
    sheet2.merge_range(10, 0, 11, 0, u"序号", cell_format = fmt1)
    for i in range(len(result_list[0])):
        sheet2.write(12+i,0,i+1,fmt9)
    
    gener_list_name = ['上库水位','上库库容','库容差值','下库水位','下库库容',
                       '毛水头','水头损失','净水头','出力','发电流量','发电时间','发电量']
    gener_list_danwei = ['m','万m3','万m3','m','万m3',
                       'm','m','m','万kW','m3/s','s','万kW.h']
    
    fmt10 = book.add_format({'font_size': 10, 'bold': True, 'border': 1, 'fg_color': '#e0fc9c'})
    fmt11 = book.add_format({'font_name':'Arial','italic': True, 'border': 1, 'fg_color': '#FFE4E1'})
    for i in range(len(gener_list_name)):
        sheet2.write(10,i+1,gener_list_name[i],fmt10)
    for i in range(len(gener_list_danwei)):
        sheet2.write(11,i+1,gener_list_danwei[i],fmt11)
    for i in range(len(result_list[0])):
        sheet2.write(12+i,1,result_list[0][i],fmt5)
    for i in range(len(result_list[1])):
        sheet2.write(12+i,2,result_list[1][i],fmt4)
    for i in range(len(result_list[2])):
        sheet2.write(12+i,3,result_list[2][i],fmt4)
    for i in range(len(result_list[3])):
        sheet2.write(12+i,4,result_list[3][i],fmt5)
    for i in range(len(result_list[4])):
        sheet2.write(12+i,5,result_list[4][i],fmt4)
    for i in range(len(result_list[5])):
        sheet2.write(12+i,6,result_list[5][i],fmt4)
    for i in range(len(result_list[6])):
        sheet2.write(12+i,7,result_list[6][i],fmt4)
    for i in range(len(result_list[7])):
        sheet2.write(12+i,8,result_list[7][i],fmt4)
    for i in range(len(result_list[8])):
        sheet2.write(12+i,9,result_list[8][i],fmt8)
    for i in range(len(result_list[9])):
        sheet2.write(12+i,10,result_list[9][i],fmt5)
    for i in range(len(result_list[10])):
        sheet2.write(12+i,11,result_list[10][i],fmt5)
    for i in range(len(result_list[11])):
        sheet2.write(12+i,12,result_list[11][i],fmt5)
    
    #sheet2.merge_range(len(result_list[0])+12,0,len(result_list[0])+12,12,"Copyright © huangzq681@gmail.com", fmt7)
    sheet2.freeze_panes(12,0)
    
    ### 抽水电能计算表
    sheet3.merge_range(0, 0, 0, 2, u"上库特征值", cell_format = fmt1)
    sheet3.merge_range(0, 3, 0, 5, u"下库特征值", cell_format = fmt1)
    sheet3.merge_range(0, 6, 0, 8, u"抽水参数", cell_format = fmt1)
    sheet3.merge_range(0, 9, 0, 12, u"水头损失", cell_format = fmt1)
    
    SK_params_names = ['正常蓄水位 (m)','死水位 (m)','正常库容 (万m3)','死库容 (万m3)','调节库容 (万m3)',
                       '正常蓄水位 (采用,m)','正常库容 (采用,万m3)','调节库容 (采用,万m3)']
    for i in range(len(SK_params_names)):
        sheet3.merge_range(1+i, 0, 1+i, 1, SK_params_names[i], fmt2)
        
    sheet3.write(1,2,SK_params[0],fmt5)
    sheet3.write(2,2,SK_params[1],fmt5)
    sheet3.write(3,2,SK_params[2],fmt4)
    sheet3.write(4,2,SK_params[3],fmt4)
    sheet3.write(5,2,SK_params[4],fmt4)
    sheet3.write(6,2,SK_params[0],fmt5)
    sheet3.write(7,2,SK_params[2],fmt4)
    sheet3.write(8,2,SK_params[4],fmt4)
    
    XK_params_names = ['正常蓄水位 (m)','死水位 (m)','正常库容 (万m3)','死库容 (万m3)','调节库容 (万m3)',
                       '正常蓄水位 (采用,m)','正常库容 (采用,万m3)','调节库容 (采用,万m3)']
    for i in range(len(XK_params_names)):
        sheet3.merge_range(1+i, 3, 1+i, 4, XK_params_names[i], fmt2)
        
    sheet3.write(1,5,XK_params[0],fmt5)
    sheet3.write(2,5,XK_params[1],fmt5)
    sheet3.write(3,5,XK_params[2],fmt4)
    sheet3.write(4,5,XK_params[3],fmt4)
    sheet3.write(5,5,XK_params[4],fmt4)
    sheet3.write(6,5,XK_norm_hgt_app,fmt5)
    sheet3.write(7,5,XK_norm_vol_app,fmt4)
    sheet3.write(8,5,XK_norm_vol_app - XK_params[3],fmt4)
    
    pump_params_names = ['装机容量 (万kW)','入力系数','发电时间 (h)','发电时间 (s)',
                          '抽水时间 (h)','抽水电量 (万kW.h)','机组台数 (台)','单机容量 (万kW)']
    for i in range(len(pump_params_names)):
        sheet3.merge_range(1+i, 6, 1+i, 7, pump_params_names[i], fmt2)
    
    sheet3.write(1,8,pump_params[0],fmt8)
    sheet3.write(2,8,pump_params[2],fmt4)
    sheet3.write(3,8,pump_params[3],fmt4)
    sheet3.write(4,8,pump_params[4],fmt8)
    sheet3.write(5,8,pump_params[5],fmt4)
    sheet3.write(6,8,pump_params[6],fmt8)
    sheet3.write(7,8,pump_params[1],fmt8)
    sheet3.write(8,8,pump_params[0]/gener_params[1],fmt5)
    

    sheet3.set_column(9,12,12)
    sheet3.merge_range(1,9,1,10,"全机抽水损失系数",fmt2)
    sheet3.merge_range(1,11,1,12,"单机抽水损失系数",fmt2)
    sheet3.merge_range(2,9,2,10,pump_loscoef_all.value,fmt6)
    sheet3.merge_range(2,11,2,12,pump_loscoef_one.value,fmt6)
    
    fmt_perc = book.add_format({'font_name':'Times New Roman','num_format':'0.00%', 'border': 1, 'font_color':'red'})
    sheet3.merge_range(7,9,8,10,"发电量/抽水电量",fmt2)
    sheet3.merge_range(7,11,8,12,gener_params[6]/pump_params[6],fmt_perc)
    ##装机容量,机组台数,入力系数,发电时间(h),发电时间(s),计算抽水时间(s),累计抽水电量
    
    head_loss_name1_pump = [u'最大毛水头 (m)',u'发电流量 (m3/s)',u'水头损失 (m)',u'最大扬程 (m)']
    head_loss_name2_pump = [u'最小毛水头 (m)',u'发电流量 (m3/s)',u'水头损失 (m)',u'最小扬程 (m)']
    head_loss_list1_pump = [head_loss_list_pump['最大毛水头'],head_loss_list_pump['最大扬程对应发电流量'],
                            head_loss_list_pump['最大扬程对应水损'],head_loss_list_pump['最大扬程']]
    head_loss_list2_pump = [head_loss_list_pump['最小毛水头'],head_loss_list_pump['最小扬程对应发电流量'],
                            head_loss_list_pump['最小扬程对应水损'],head_loss_list_pump['最小扬程']]
    
    fmt9 = book.add_format({'font_size': 10, 'bold': True, 'border': 1, 'fg_color': '#FAEBD7'})
    
    for i in range(len(head_loss_name1_pump)):
        sheet3.write(3, 9+i, head_loss_name1_pump[i], fmt9)
    for i in range(len(head_loss_name2_pump)):
        sheet3.write(5, 9+i, head_loss_name2_pump[i], fmt9)
    for i in range(len(head_loss_list1_pump)):
        sheet3.write(4, 9+i, head_loss_list1_pump[i], fmt4)
    for i in range(len(head_loss_list2_pump)):
        sheet3.write(6, 9+i, head_loss_list2_pump[i], fmt4)
        
    sheet3.merge_range(10, 0, 11, 0, u"序号", cell_format = fmt1)
    for i in range(len(result_list2[0])):
        sheet3.write(12+i,0,i+1,fmt9)
    
    pump_list_name = ['上库水位','上库库容','库容差值','下库水位','下库库容',
                       '毛水头','水头损失','净水头','出力','发电流量','发电时间','发电量']
    pump_list_danwei = ['m','万m3','万m3','m','万m3',
                       'm','m','m','万kW','m3/s','s','万kW.h']
    
    fmt10 = book.add_format({'font_size': 10, 'bold': True, 'border': 1, 'fg_color': '#e0fc9c'})
    fmt11 = book.add_format({'font_name':'Arial','italic': True, 'border': 1, 'fg_color': '#FFE4E1'})
    for i in range(len(pump_list_name)):
        sheet3.write(10,i+1,pump_list_name[i],fmt10)
    for i in range(len(pump_list_danwei)):
        sheet3.write(11,i+1,pump_list_danwei[i],fmt11)
    for i in range(len(result_list2[0])):
        sheet3.write(12+i,1,result_list2[0][i],fmt5)
    for i in range(len(result_list2[1])):
        sheet3.write(12+i,2,result_list2[1][i],fmt4)
    for i in range(len(result_list2[2])):
        sheet3.write(12+i,3,result_list2[2][i],fmt4)
    for i in range(len(result_list2[3])):
        sheet3.write(12+i,4,result_list2[3][i],fmt5)
    for i in range(len(result_list2[4])):
        sheet3.write(12+i,5,result_list2[4][i],fmt4)
    for i in range(len(result_list2[5])):
        sheet3.write(12+i,6,result_list2[5][i],fmt4)
    for i in range(len(result_list2[6])):
        sheet3.write(12+i,7,result_list2[6][i],fmt4)
    for i in range(len(result_list2[7])):
        sheet3.write(12+i,8,result_list2[7][i],fmt4)
    for i in range(len(result_list2[8])):
        sheet3.write(12+i,9,result_list2[8][i],fmt8)
    for i in range(len(result_list2[9])):
        sheet3.write(12+i,10,result_list2[9][i],fmt5)
    for i in range(len(result_list2[10])):
        sheet3.write(12+i,11,result_list2[10][i],fmt5)
    for i in range(len(result_list2[11])):
        sheet3.write(12+i,12,result_list2[11][i],fmt5)
    
    # sheet3.merge_range(len(result_list2[0])+12,0,len(result_list2[0])+12,12,"Copyright © huangzq681@gmail.com", fmt7)
    sheet3.freeze_panes(12,0)
    
    ### 蓄能量曲线
    sheet4.merge_range(0, 0, 0, 2, u"上库特征值", cell_format = fmt1)
    sheet4.merge_range(0, 3, 0, 5, u"下库特征值", cell_format = fmt1)
    sheet4.merge_range(0, 6, 0, 8, u"发电参数", cell_format = fmt1)
    
    SK_params_names = ['正常蓄水位 (m)','死水位 (m)','正常库容 (万m3)','死库容 (万m3)','调节库容 (万m3)',
                       '正常蓄水位 (采用,m)','正常库容 (采用,万m3)','调节库容 (采用,万m3)']
    for i in range(len(SK_params_names)):
        sheet4.merge_range(1+i, 0, 1+i, 1, SK_params_names[i], fmt2)
        
    sheet4.write(1,2,SK_params[0],fmt5)
    sheet4.write(2,2,SK_params[1],fmt5)
    sheet4.write(3,2,SK_params[2],fmt4)
    sheet4.write(4,2,SK_params[3],fmt4)
    sheet4.write(5,2,SK_params[4],fmt4)
    sheet4.write(6,2,SK_params[0],fmt5)
    sheet4.write(7,2,SK_params[2],fmt4)
    sheet4.write(8,2,SK_params[4],fmt4)
    
    XK_params_names = ['正常蓄水位 (m)','死水位 (m)','正常库容 (万m3)','死库容 (万m3)','调节库容 (万m3)',
                       '正常蓄水位 (采用,m)','正常库容 (采用,万m3)','调节库容 (采用,万m3)']
    for i in range(len(XK_params_names)):
        sheet4.merge_range(1+i, 3, 1+i, 4, XK_params_names[i], fmt2)
        
    sheet4.write(1,5,XK_params[0],fmt5)
    sheet4.write(2,5,XK_params[1],fmt5)
    sheet4.write(3,5,XK_params[2],fmt4)
    sheet4.write(4,5,XK_params[3],fmt4)
    sheet4.write(5,5,XK_params[4],fmt4)
    sheet4.write(6,5,XK_norm_hgt_app,fmt5)
    sheet4.write(7,5,XK_norm_vol_app,fmt4)
    sheet4.write(8,5,XK_norm_vol_app - XK_params[3],fmt4)
    
    gener_params_names = ['装机容量 (万kW)','出力系数','发电时间 (h)','发电时间 (s)',
                          '计算时间 (s)','发电量 (万kW.h)','机组台数 (台)','单机容量 (万kW)']
    for i in range(len(gener_params_names)):
        sheet4.merge_range(1+i, 6, 1+i, 7, gener_params_names[i], fmt2)
    ## 装机容量，机组台数，出力系数，发电时间h，发电时间s,计算发电时间s,累计发电量
        
    sheet4.write(1,8,gener_params[0],fmt8)
    sheet4.write(2,8,gener_params[2],fmt4)
    sheet4.write(3,8,gener_params[3],fmt4)
    sheet4.write(4,8,gener_params[4],fmt8)
    sheet4.write(5,8,gener_params[5],fmt8)
    sheet4.write(6,8,gener_params[6],fmt8)
    sheet4.write(7,8,gener_params[1],fmt8)
    sheet4.write(8,8,gener_params[0]/gener_params[1],fmt5)
    
    sheet4.merge_range(10, 0, 11, 0, u"序号", cell_format = fmt1)
    for i in range(len(result_list[0])):
        sheet4.write(12+i,0,i+1,fmt9)
    
    gener_curve_name = ['上库水位','上库库容','下库水位','下库库容','出力','发电时间','发电量']
    gener_curve_danwei = ['m','万m3','m','万m3','万kW','s','万kW.h']
    
    if cal_xnqx.get() == 1:
        for i in range(len(gener_curve_name)):
            sheet4.write(10,i+1,gener_curve_name[i],fmt10)
        for i in range(len(gener_curve_danwei)):
            sheet4.write(11,i+1,gener_curve_danwei[i],fmt11)
        for i in range(len(result_list2[0])):
            sheet4.write(12+i,1,energy_storage_curve1[0][i],fmt5)
        for i in range(len(result_list2[1])):
            sheet4.write(12+i,2,energy_storage_curve1[1][i],fmt4)
        for i in range(len(result_list2[2])):
            sheet4.write(12+i,3,energy_storage_curve1[2][i],fmt5)
        for i in range(len(result_list2[3])):
            sheet4.write(12+i,4,energy_storage_curve1[3][i],fmt4)
        for i in range(len(result_list2[4])):
            sheet4.write(12+i,5,energy_storage_curve1[4][i],fmt8)
        for i in range(len(result_list2[5])):
            sheet4.write(12+i,6,energy_storage_curve1[5][i],fmt4)
        for i in range(len(result_list2[6])):
            sheet4.write(12+i,7,energy_storage_curve1[6][i],fmt5)
        
        # sheet4.merge_range(len(result_list2[0])+12,0,len(result_list2[0])+12,7,"Copyright © huangzq681@gmail.com", fmt7)
        sheet4.freeze_panes(15,0)
        
        gener_curve_chart = book.add_chart({'type':'line'}) 
        gener_curve_chart.add_series(
            {
             'name':'=蓄能量曲线!$H$11',
             'categories':'=蓄能量曲线!$B$13:$B$'+ str(13+len(result_list2[0])-1),
             'values':'=蓄能量曲线!$H$13:$H$'+ str(13+len(result_list2[0])-1),
             'line':{'color':'red','width':1}
             })
        
        gener_curve_chart.set_title({'name':'蓄能量曲线图',
                                     'name_font': {'name':'微软雅黑','bold': False, 'italic': False,'size': 12}})
        gener_curve_chart.set_x_axis({'name':'水位(m)','num_format':'0',
                                      'name_font': {'name':'微软雅黑','bold': False, 'italic': False,'size': 10},
                                      'min':str(round(SK_params[0]+1,0)), 
                                      'max':str(round(SK_params[1]-1,0)),
                                      'interval_tick': 2,
                                      'interval_unit': 20,
                                      #'minor_unit': 2, 'major_unit': 10,
                                      'major_gridlines': {
                                          'visible': True,
                                          'line': {'width': 0.5}}})
        gener_curve_chart.set_y_axis({'name':'发电量(万kW.h)','num_format':'0',
                                      'name_font': {'name':'微软雅黑','bold': False, 'italic': False,'size': 10}})
        gener_curve_chart.set_legend({
            'layout': {
                'x':      0.7,
                'y':      0.6,
                'width':  0.25,
                'height': 0.25,
            }
        })
        
        gener_curve_chart.set_style(1)
        gener_curve_chart.set_plotarea({
            'layout': {
                'x':      0.10,
                'y':      0.12,
                'width':  0.85,
                'height': 0.70,
            }
        })
        sheet4.insert_chart('J1',gener_curve_chart, {'x_offset':25,'y_offset':10})
    else:
        pass
    
    ### 上库库容曲线
    sheet5.merge_range(0, 0, 1, 0, u"序号", fmt1)
    sheet5.write(0, 1, u"高程", fmt1)
    sheet5.write(0, 2, u"库容", fmt1)
    sheet5.write(1, 1, u"m", fmt11)
    sheet5.write(1, 2, u"万m3", fmt11)
    
    for i in range(len(SK_hgts)):
        sheet5.write(i+2,0,i,fmt2)
        sheet5.write(i+2,1,SK_hgts[i],fmt5)
        sheet5.write(i+2,2,SK_vols[i],fmt5)
    
    sk_curve_chart = book.add_chart({'type':'line'}) 
    sk_curve_chart.add_series(
        {
         'name':'=上库库容曲线!$C$1',
         'categories':'=上库库容曲线!$B$3:$B$'+ str(3+len(SK_hgts)-1),
         'values':'=上库库容曲线!$C$3:$C$'+ str(3+len(SK_hgts)-1),
         'line':{'color':'red','width':1}
         })
    
    sk_curve_chart.set_title({'name':'上库库容曲线图',
                                 'name_font': {'name':'微软雅黑','bold': False, 'italic': False,'size': 12}})
    sk_curve_chart.set_x_axis({'name':'水位(m)','num_format':'0',
                                  'name_font': {'name':'微软雅黑','bold': False, 'italic': False,'size': 10},
                                  'min':str(round(SK_params[0]+1,0)), 
                                  'max':str(round(SK_params[1]-1,0)),
                                  'interval_tick': 5,
                                  'interval_unit': 5,
                                  'major_gridlines': {
                                      'visible': True,
                                      'line': {'width': 0.5}}})
    sk_curve_chart.set_y_axis({'name':'库容(万m3)','num_format':'0',
                                  'name_font': {'name':'微软雅黑','bold': False, 'italic': False,'size': 10}})
    sk_curve_chart.set_legend({
        'layout': {
            'x':      0.7,
            'y':      0.6,
            'width':  0.25,
            'height': 0.25,
        }
    })
    
    sk_curve_chart.set_style(1)
    sk_curve_chart.set_plotarea({
        'layout': {
            'x':      0.15,
            'y':      0.12,
            'width':  0.80,
            'height': 0.72,
        }
    })
    sheet5.insert_chart('D1',sk_curve_chart, {'x_offset':25,'y_offset':10})
    
    ### 下库库容曲线
    sheet6.merge_range(0, 0, 1, 0, u"序号", fmt1)
    sheet6.write(0, 1, u"高程", fmt1)
    sheet6.write(0, 2, u"库容", fmt1)
    sheet6.write(1, 1, u"m", fmt11)
    sheet6.write(1, 2, u"万m3", fmt11)
    
    for i in range(len(XK_hgts)):
        sheet6.write(i+2,0,i,fmt2)
        sheet6.write(i+2,1,XK_hgts[i],fmt5)
        sheet6.write(i+2,2,XK_vols[i],fmt5)
    
    xk_curve_chart = book.add_chart({'type':'line'}) 
    xk_curve_chart.add_series(
        {
         'name':'=下库库容曲线!$C$1',
         'categories':'=下库库容曲线!$B$3:$B$'+ str(3+len(XK_hgts)-1),
         'values':'=下库库容曲线!$C$3:$C$'+ str(3+len(XK_hgts)-1),
         'line':{'color':'red','width':1}
         })
    
    xk_curve_chart.set_title({'name':'下库库容曲线图',
                                 'name_font': {'name':'微软雅黑','bold': False, 'italic': False,'size': 12}})
    xk_curve_chart.set_x_axis({'name':'水位(m)','num_format':'0',
                                  'name_font': {'name':'微软雅黑','bold': False, 'italic': False,'size': 10},
                                  'min':str(round(XK_params[0]+1,0)), 
                                  'max':str(round(XK_params[1]-1,0)),
                                  'interval_tick': 5,
                                  'interval_unit': 5,
                                  'major_gridlines': {
                                      'visible': True,
                                      'line': {'width': 0.5}}})
    xk_curve_chart.set_y_axis({'name':'库容(万m3)','num_format':'0',
                               'name_font': {'name':'微软雅黑','bold': False, 'italic': False,'size': 10}})
    xk_curve_chart.set_legend({
        'layout': {
            'x':      0.7,
            'y':      0.6,
            'width':  0.25,
            'height': 0.25,
        }
    })
    
    xk_curve_chart.set_style(1)
    xk_curve_chart.set_plotarea({
        'layout': {
            'x':      0.15,
            'y':      0.12,
            'width':  0.80,
            'height': 0.72,
        }
    })
    sheet6.insert_chart('D1',xk_curve_chart, {'x_offset':25,'y_offset':10})
    
    book.close()
    print("写文档完成！文档目录为："+res_filepath.get())



