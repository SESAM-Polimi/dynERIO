# -*- coding: utf-8 -*-
"""
Created on Thu Jun 23 08:44:13 2022

@author: loren
"""

import mario

#%% parse hybrid database
path_exiobase_hybrid = r"C:/Users/loren/OneDrive - Politecnico di Milano/Politecnico di Milano/Research/Database/Exiobase/3.3.18/EXIOBASE_3_3_18_hsut_2011.zip"
world = mario.parse_exiobase_sut(path_exiobase_hybrid,units="hybrid",version="3.3.18",extensions="all")

#%% convert commodities units
path_unit_conv = r"support_excels\unit_conv.xlsx"
# world.get_unit_conversion_excel(path_unit_conv)
world.convert_units(path_unit_conv,items=["Commodity"])

#%% aggregate database
path_aggregate_sut = r"support_excels\aggregate_sut.xlsx"
# world.get_aggregation_excel(path_aggregate_sut)
world.aggregate(path_aggregate_sut)

#%% SUT to IOT prod-by-prod transformation
world.sut_to_iot(method="B")

#%% parse economic database
path_exio_pxp = r"C:\Users\loren\OneDrive - Politecnico di Milano\Politecnico di Milano\Research\Database\Exiobase\3.8.2\economic\IOT\pxp\IOT_2011_pxp.zip"
iot = mario.parse_exiobase_3(path_exio_pxp)

#%% aggregate database
path_aggregate_iot = r"support_excels\aggregate_iot.xlsx"
# world.get_aggregation_excel(path_aggregate_sut)
iot.aggregate(path_aggregate_iot)

#%%
world.matrices['baseline']['V'].index = iot.matrices['baseline']['V'].index
world.matrices['baseline']['E'].loc["Employment",:] = 0

#%% Adding value added and extension matrices from iot tables into sut table + Adjusting units
for region in world.get_index("Region"):
    world.matrices['baseline']['V'].loc[:,(region,slice(None),slice(None))] = iot.matrices['baseline']['V'].loc[:,(region,slice(None),slice(None))]
    for satellite in iot.get_index("Satellite account"):
        world.matrices['baseline']['E'].loc[:,(region,slice(None),slice(None))] = iot.matrices['baseline']['E'].loc[:,(region,slice(None),slice(None))]

#%%
# for region_row in world.get_index("Region"):
#     for region_col in world.get_index("Region"):
#         world.matrices['baseline']['Z'].loc[(region_row,slice(None),"goods"),(region_col,slice(None),"goods")]    = iot.matrices['baseline']['Z'].loc[(region_row,slice(None),"goods"),(region_col,slice(None),"goods")].values
#         world.matrices['baseline']['Z'].loc[(region_row,slice(None),"goods"),(region_col,slice(None),"services")] = iot.matrices['baseline']['Z'].loc[(region_row,slice(None),"goods"),(region_col,slice(None),"services")].values
#         world.matrices['baseline']['Z'].loc[(region_row,slice(None),"goods"),(region_col,slice(None),"steam and hot water")] = iot.matrices['baseline']['Z'].loc[(region_row,slice(None),"goods"),(region_col,slice(None),"steam and hot water")].values
    
# #%% aggregate database
# path_aggregate_fin = r"support_excels\aggregate_fin.xlsx"
# world.get_aggregation_excel(path_aggregate_sut)
# world.aggregate(path_aggregate_fin)

world.to_excel(r"support_excels/hybrid.xlsx")
iot.to_excel(r"support_excels/economic.xlsx")
