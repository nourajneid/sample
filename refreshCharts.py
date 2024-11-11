import sys
import os
import time

if sys.platform == "win32":
    import win32com.client as win32
    import pythoncom
    from multiprocessing import Process
    import subprocess

vba_code = """
' if run by the user manually then only this function is used
Sub RunUpdateChartDataAndRedraw()
    ' Call the main macro, passing the active presentation
    UpdateChartDataAndRedraw
End Sub

Sub UpdateChartDataAndRedraw(Optional pres As Presentation = Nothing)
    Dim targetPres As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim groupStack As Collection
    Dim i As Integer

    ' Check if a presentation was passed. If not, use ActivePresentation
    If pres Is Nothing Then
        Set targetPres = ActivePresentation
    Else
        Set targetPres = pres
    End If

    ' Loop through all slides in the target presentation
    For Each sld In targetPres.Slides
        ' Iterate through shapes in the current slide
        For Each shp In sld.Shapes
            ' If shape is a chart, refresh it
            If shp.HasChart Then
                RefreshChart shp
            ' If shape is a group, add its items to the stack for iterative processing
            ElseIf shp.Type = msoGroup Then
                Set groupStack = New Collection
                For i = 1 To shp.GroupItems.Count
                    groupStack.Add shp.GroupItems(i)
                Next i

                ' Process group items iteratively
                While groupStack.Count > 0
                    Set shp = groupStack(1)
                    groupStack.Remove 1

                    If shp.HasChart Then
                        RefreshChart shp
                    ElseIf shp.Type = msoGroup Then
                        ' Add nested group items to the stack
                        For i = 1 To shp.GroupItems.Count
                            groupStack.Add shp.GroupItems(i)
                        Next i
                    End If
                Wend
            End If
        Next shp
    Next sld
End Sub

Sub RefreshChart(shp As Shape)
    Dim cht As Chart
    Dim wb As Object ' Using late binding to avoid library dependency issues

    ' Try to set the chart object
    On Error Resume Next
    Set cht = shp.Chart
    On Error GoTo 0

    ' If the chart object is valid and chart data is not linked, refresh it
    If Not cht Is Nothing Then
        If Not cht.ChartData Is Nothing Then
            If Not cht.ChartData.IsLinked Then
                ' Use late binding to access workbook to avoid dependency on Excel library
                On Error Resume Next
                Set wb = cht.ChartData.Workbook
                If Not wb Is Nothing Then
                    wb.Name ' Access the workbook to force refresh
                End If
                On Error GoTo 0
            End If
        End If
    End If
End Sub
"""

LOCK_FILE = "powerpoint_process.lock"

def kill_powerpoint():
    start_time = time.time()
    try:
        # Kill any running PowerPoint process using PowerShell
        subprocess.run(["powershell", "-Command", "Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue"], check=True)
    except subprocess.CalledProcessError:
        # Ignore error if PowerPoint process is not found
        print("PowerPoint process not found, no need to kill.")
    end_time = time.time()
    print(f"PowerPoint kill process took {end_time - start_time:.2f} seconds")

def refreshCharts(output_ppt_file):
    total_start_time = time.time()
    print("in refresh")
    
    # Kill any running instance of PowerPoint before starting the process
    kill_powerpoint()

    pythoncom.CoInitialize()  # Initialize COM in the current thread
    init_time = time.time()
    print(f"COM initialization took {init_time - total_start_time:.2f} seconds")

    try:
        # Start PowerPoint and make it visible
        start_powerpoint_time = time.time()
        powerpoint = win32.DispatchEx("PowerPoint.Application")
        open_powerpoint_time = time.time()
        print(f"Starting PowerPoint took {open_powerpoint_time - start_powerpoint_time:.2f} seconds")

        # Open the PowerPoint presentation
        open_presentation_time = time.time()
        presentation = powerpoint.Presentations.Open(output_ppt_file, WithWindow=True)
        presentation_opened_time = time.time()
        print(f"Opening the PowerPoint presentation took {presentation_opened_time - open_presentation_time:.2f} seconds")

        # Add and run the VBA macro, passing the presentation
        add_module_time = time.time()
        module = presentation.VBProject.VBComponents.Add(1)  # Add a new module
        module.CodeModule.AddFromString(vba_code)
        vba_added_time = time.time()
        print(f"Adding VBA macro took {vba_added_time - add_module_time:.2f} seconds")

        # Run the macro and pass the presentation to it
        run_macro_time = time.time()
        #viewtype = powerpoint.ActiveWindow.ViewType
        #powerpoint.ActiveWindow.ViewType = 1  # ppViewNormal to ensure the presentation is updated
        powerpoint.Run(f"{module.Name}.UpdateChartDataAndRedraw", presentation)
        macro_run_time = time.time()
        print(f"Running VBA macro took {macro_run_time - run_macro_time:.2f} seconds")

        # Remove the VBA module after running the macro
        remove_module_time = time.time()
        try:
            presentation.VBProject.VBComponents.Remove(module)
        except Exception as e:
            print(f"An error occurred while removing the VBA module: {e}")
        module_removed_time = time.time()
        print(f"Removing VBA module took {module_removed_time - remove_module_time:.2f} seconds")

        # Save the presentation after running the macro and removing the module
        save_presentation_time = time.time()
        file_extension = os.path.splitext(output_ppt_file)[1].lower()

        # Determine the correct format type based on the file extension
        if file_extension == ".pptx":
            format_type = 24  # ppSaveAsOpenXMLPresentation
        elif file_extension == ".pptm":
            format_type = 25  # ppSaveAsOpenXMLMacroEnabledPresentation
        else:
            format_type = 32  # ppSaveAsDefault, fallback option

        print("before path")
        absolute_path = os.path.abspath(output_ppt_file)
        print("absolute path")
        #powerpoint.ActiveWindow.ViewType = viewtype
        time.sleep(2)  # Wait for 2 seconds before saving to ensure all updates are applied
        presentation.SaveAs(absolute_path, format_type)
        print(f"PowerPoint saved at: {absolute_path}")
        presentation_saved_time = time.time()
        print(f"Saving the PowerPoint presentation took {presentation_saved_time - save_presentation_time:.2f} seconds")
    except Exception as e:
        print(f"An error occurred in PowerPoint VBA execution: {e}")

    finally:
        # Ensure PowerPoint is properly released
        quit_time = time.time()
        powerpoint.Quit()
        pythoncom.CoUninitialize()
        final_time = time.time()
        print(f"Quitting PowerPoint took {final_time - quit_time:.2f} seconds")
        print(f"Total execution time: {final_time - total_start_time:.2f} seconds")
