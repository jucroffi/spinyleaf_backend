import streamlit as st
import sys
import os

import json
import pathlib


from ladybug.color import Color
import honeybee_vtk
from honeybee_vtk.scene import Scene
from honeybee_vtk.camera import Camera
from honeybee_vtk.model import DisplayMode
from honeybee_vtk.model import (HBModel,
                                Model as VTKModel,
                                SensorGridOptions,
                                DisplayMode)
from honeybee_vtk.legend_parameter import ColorSets

from honeybee_radiance.view import View

from streamlit_vtkjs import st_vtkjs



def get_config(study_name, results_folder, unit, d_range, cs):

    cfg = {
        "data": [
            {
                "identifier": study_name,
                "object_type": "grid",
                "unit": unit[study_name],
                "path": results_folder.as_posix(),
                "hide": False,
                "legend_parameters": {
                        "hide_legend": False,
                        "min": d_range[study_name][0],
                        "max": d_range[study_name][1],
                        "color_set": cs[study_name],
                        "label_parameters": {
                            "color": [0, 0, 0],
                            "size": 0,
                            "bold": True
                        }
                }
            }
        ]
    }

    
    config_file = results_folder.joinpath("config.json")
    with open(config_file, "w") as f:
        json.dump(cfg, f, indent=2)

    return config_file




def color_vtkjs_from_results(model, results_folder, study_name):
    
    vtk_model = VTKModel(model, SensorGridOptions.Mesh)

    cs = {'Horizontal_Views': 'view_study',
          'Horizontal_Mean' : 'heat_sensation',
          'Outdoors_Views_Satisfaction' : 'shade_benefit_harm',
          'Sky_Views': 'view_study',
          'Sky_Mean' : 'cold_sensation',
          'Sky_Views_Satisfaction' : 'shade_benefit_harm',
          'Green_Views': 'view_study',
          'Green_Mean' : 'peak_load_balance',
          'Green_Views_Satisfaction' : 'shade_benefit_harm',
          "Balcony_Areas": 'shadow_study',
          "Balcony_Percentage": 'shadow_study',
          "Access_to_Green_Satisfaction": 'shade_benefit_harm',
          "Areas": 'annual_comfort',
          "Occupancy_Rate": 'energy_balance',
          "Space_Size_Satisfaction" : 'shade_benefit_harm',
          "Extreme_Hot_Week_Temp": "nuanced",
          "Ext_Hot_Thermal_Sensation": "thermal_comfort",
          "Ext_Hot_Thermal_Satisfaction" : 'shade_benefit_harm',
          "Extreme_Cold_Week_Temp": 'nuanced',
          "Ext_Cold_Thermal_Sensation": 'thermal_comfort',
          "Ext_Cold_Thermal_Satisfaction": 'shade_benefit_harm',
          "Daylight_Autonomy": "ecotect",
          "DA_mean": 'ecotect',
          "Daylight_Satisfaction": 'shade_benefit_harm',
          "Useful_Daylight_Illuminance": "ecotect",
          "UDI_mean": "ecotect",
          "Glare_Autonomy": "glare_study",
          "GA_mean": "glare_study",
          "CO2_Levels": 'black_to_white',
          "Relative_Humidity": 'cloud_cover',
          "Air_Quality_Satisfaction": 'shade_benefit_harm',
          "Delight_Satisfaction": 'benefit_harm',
          "Sound_Levels": 'blue_green_red',
          "Sound_Level_Satisfaction": 'shade_benefit_harm',
          "Comfort_Satisfaction": 'benefit_harm',
          "Social_Green_Areas": 'peak_load_balance',
          "Social_Green_Area_Occupants": 'peak_load_balance',
          "Social_green_satisfaction":'shade_benefit_harm',
          "Social_Total_Areas": 'shadow_study',
          "Social_Area_Occupants": 'shadow_study',
          "Social_Amount_Satisfaction": 'shade_benefit_harm',
          "Social_Levels_Available": 'annual_comfort',
          "Weighted_Distribution_Social_Spaces": 'annual_comfort',
          "Social_Distribution_Satisfaction": 'shade_benefit_harm',
          "Social_Satisfaction": 'benefit_harm',
          "Wellbeing_Fostered_by_Design": 'benefit_harm'}

    
    d_range = {'Horizontal_Views': [0, 50],
               'Horizontal_Mean' : [0, 30],
               'Outdoors_Views_Satisfaction' : [0, 2],
               'Sky_Views': [0, 50],
               'Sky_Mean' : [0, 30],
               'Sky_Views_Satisfaction' : [0, 2],
               'Green_Views': [0, 50],
               'Green_Mean' : [0, 20],
               'Green_Views_Satisfaction' : [0, 2],
                "Balcony_Areas": [0,15],
                "Balcony_Percentage": [0,15],
                "Access_to_Green_Satisfaction":[0, 2],
                "Areas": [30,120],
                "Occupancy_Rate": [0.016,0.026],
                "Space_Size_Satisfaction":[0, 2],
                "Extreme_Hot_Week_Temp": [24, 30],
                "Ext_Hot_Thermal_Sensation": [-2,2],
                "Ext_Hot_Thermal_Satisfaction" : [0, 2],
                "Extreme_Cold_Week_Temp": [15, 24],
                "Ext_Cold_Thermal_Sensation": [-2,2],
                "Ext_Cold_Thermal_Satisfaction": [0, 2],
                "Daylight_Autonomy": [0,100],
                "DA_mean": [40,80],
                "Daylight_Satisfaction": [0, 1],
                "Useful_Daylight_Illuminance": [0,100],
                "UDI_mean":[40,80],
                "Glare_Autonomy": [0,100],
                "GA_mean": [40,80],
                "CO2_Levels": [400,1000],
                "Relative_Humidity": [40,70],
                "Air_Quality_Satisfaction":[0, 2],
                "Delight_Satisfaction":[0,2],
                "Sound_Levels": [30,50],
                "Sound_Level_Satisfaction": [0,1],
                "Comfort_Satisfaction":[0,2],
                "Social_Green_Areas": [100,200],
                "Social_Green_Area_Occupants": [0,20],
                "Social_green_satisfaction":[0,2],
                "Social_Total_Areas": [50,500],
                "Social_Area_Occupants": [0,20],
                "Social_Amount_Satisfaction":[0,2],
                "Social_Levels_Available": [0,4],
                "Weighted_Distribution_Social_Spaces": [0,1],
                "Social_Distribution_Satisfaction":[0,2],
                "Social_Satisfaction": [0,2],
                "Wellbeing_Fostered_by_Design": [0,6]}
    
    
    unit = {'Horizontal_Views': "%",
            'Horizontal_Mean' : "% mean",
            'Outdoors_Views_Satisfaction' : "Satisfaction", 
            'Sky_Views': "%",
            'Sky_Mean' : "% mean",
            'Sky_Views_Satisfaction' : "Satisfaction", 
            'Green_Views': "%",
            'Green_Mean' : "% mean",
            'Green_Views_Satisfaction' : "Satisfaction",
            "Balcony_Areas": 'Area m2',
            "Balcony_Percentage": '%',
            "Access_to_Green_Satisfaction": "Satisfaction",
            "Areas": 'Area m2',
            "Occupancy_Rate": 'ppl/m2',
            "Space_Size_Satisfaction" :"Satisfaction",
            "Extreme_Hot_Week_Temp": "C",
            "Ext_Hot_Thermal_Sensation": "Cold / Warm",
            "Ext_Hot_Thermal_Satisfaction" : 'Satisfaction',
            "Extreme_Cold_Week_Temp": "C",
            "Ext_Cold_Thermal_Sensation": "Cold / Warm",
            "Ext_Cold_Thermal_Satisfaction": 'Satisfaction',
            "Daylight_Autonomy": "DA %",
            "DA_mean": "% mean",
            "Daylight_Satisfaction": 'Satisfaction',
            "Useful_Daylight_Illuminance": "UDI %",
            "UDI_mean": "% mean",
            "Glare_Autonomy": "GA %",
            "GA_mean": "% mean",
            "CO2_Levels": 'ppm',
            "Relative_Humidity": '%',
            "Air_Quality_Satisfaction":'Satisfaction',
            "Delight_Satisfaction": 'Satisfaction',
            "Sound_Levels": 'dB',
            "Sound_Level_Satisfaction":'Satisfaction',
            "Comfort_Satisfaction":'Satisfaction',
            "Social_Green_Areas": 'Area m2',
            "Social_Green_Area_Occupants": 'm2/occupant',
            "Social_green_satisfaction":'Satisfaction',
            "Social_Total_Areas": 'm2/occupant',
            "Social_Area_Occupants": 'm2/occupant',
            "Social_Amount_Satisfaction":'Satisfaction',
            "Social_Levels_Available": 'Number of Social Levels Available',
            "Weighted_Distribution_Social_Spaces": 'wdss',
            "Social_Distribution_Satisfaction": 'Satisfaction',
            "Social_Satisfaction": 'Satisfaction',
            "Wellbeing_Fostered_by_Design": 'Satisfaction'}
    
    
    vtk_model.sensor_grids.display_mode = DisplayMode.SurfaceWithEdges
    vtk_model.shades.display_mode = DisplayMode.Surface
    vtk_model.shades.color = Color(130, 130, 130, 200)
    vtk_model.walls.display_mode = DisplayMode.Wireframe
    vtk_model.walls.color = Color(0, 0, 0, 255)
    vtk_model.floors.display_mode = DisplayMode.Wireframe
    vtk_model.floors.color = Color(0, 0, 0, 0)
    vtk_model.roof_ceilings.display_mode = DisplayMode.Wireframe
    vtk_model.roof_ceilings.color = Color(0, 0, 0, 0)


    config_file = get_config(study_name, results_folder, unit, d_range, cs)

    vtk_model.to_vtkjs(folder=results_folder,name=study_name, 
                       config=config_file.as_posix(),
                       model_display_mode=DisplayMode.Wireframe)
    

    
def get_vtkjs(model, output_folder, name):

    view = View(identifier='view_test', position=(0, 0, 100), direction=(5,20,2), up_vector=None, type='v', h_size=40, v_size=40)
    model.properties.radiance.add_view(view)
    vtk_model = VTKModel(model, SensorGridOptions.Sensors)
    vtk_model.shades.display_mode = DisplayMode.SurfaceWithEdges
    vtk_model.shades.color = Color(130, 130, 130, 180)
    vtk_model.walls.display_mode = DisplayMode.Surface
    vtk_model.walls.color = Color(250, 250, 250, 255)
    vtk_model.roof_ceilings.display_mode = DisplayMode.Surface
    vtk_model.roof_ceilings.color = Color(109, 148, 105, 255)


    vtk_model.to_vtkjs(folder=output_folder, name = name,
                       model_display_mode=DisplayMode.Surface)
    model.to_hbjson(folder=output_folder, name = name)


def get_views(name, output_folder, height):

    key = f'{name}_{0}'
    read_folder = pathlib.Path('data', output_folder, f'{name}.vtkjs').read_bytes()
    height = f'{height}px'
    views = st_vtkjs(content=read_folder, key=key,style={'height': height}, subscribe=False)
    return views


def view_study(view_hb_model, study_dics, study_names, labels, v_height):

    for n, column in enumerate(st.columns(len(study_names))):
        color_vtkjs_from_results(view_hb_model, study_dics[study_names[n]], study_name=study_names[n])
        with column:
            st.success(labels[n])
            get_views(study_names[n], study_dics[study_names[n]], v_height)


  
