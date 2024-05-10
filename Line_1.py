# -*- coding: utf-8 -*-
"""
Created on Thu May  9 13:05:47 2024

@author: Santanu
"""

import win32com.client

def draw_line(start_point, end_point):
    # Connect to SolidWorks
    swApp = win32com.client.Dispatch("SldWorks.Application")

    # Check if SolidWorks is already running
    if not swApp:
        print("SolidWorks is not running")
        return

    # Try to create a new document
    try:
        swModel = swApp.NewDocument("Part", 0, 0, 0)
    except Exception as e:
        print("Failed to create a new document:", e)
        return

    if not swModel:
        print("Failed to create a new part document")
        return

    # Access the active sketch
    swSketchManager = swModel.SketchManager
    if not swSketchManager:
        print("Failed to access the sketch manager")
        return

    # Insert a sketch
    success = swSketchManager.InsertSketch(True)
    if not success:
        print("Failed to insert a sketch")
        return

    # Create a new sketch segment (line)
    swSketchManager.CreateLine(start_point[0], start_point[1], start_point[2], end_point[0], end_point[1], end_point[2])

    # Exit the sketch
    swModel.ClearSelection2(True)
    swModel.EditRebuild3()

if __name__ == "__main__":
    # Define the start and end points of the line
    start_point = (0, 0, 0)  # (x, y, z)
    end_point = (10, 0, 0)   # (x, y, z)

    # Draw the line
    draw_line(start_point, end_point)
