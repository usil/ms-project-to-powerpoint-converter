import os
import re
import win32com.client as w3c
from pathlib import Path
from functions.utils import getParentTask, getChildTasks, getChildTasksByParent

def fileExists(sourceFolder):
    sourcePptx = None
    sourceMpp = None
    
    with os.scandir(sourceFolder) as files:
        for file in files: 
            if file.name.endswith('.pptx'):
                sourcePptx = file
            elif file.name.endswith('.mpp'):
                sourceMpp = file
                
    return sourcePptx, sourceMpp

def getTasksByName(baseNode: str, file):
    mpp: w3c.CDispatch = w3c.Dispatch("MSProject.Application")
    mpp.FileOpen(file.path)
    masterProject = mpp.ActiveProject
      
    tasksMasterProject = masterProject.Tasks  
    taskParent = None    
    listBaseNode = baseNode.split('>')
    
    print('--------------------------------------')
    print(f"Validate file MS Project {file.name}")
    print('--------------------------------------')
    
    if len(baseNode.strip()) == 0:
        print('Base node is empty')
        print('--------------------------------------')
        return None
  
    indexBaseNode = 0
    for node in listBaseNode:
        task = next((x for x in tasksMasterProject 
                        if x is not None 
                        and x.Name is not None 
                        and x.Name.strip() == node.strip()), 
                    None)
        if task is not None and len(task.Name.strip()) > 0:
            if indexBaseNode == 0 or (getParentTask(task) is not None 
                                        and getParentTask(task).strip() == listBaseNode[indexBaseNode - 1].strip()):
                taskParent = task
                if indexBaseNode == len(listBaseNode) - 1:
                    tasksMasterProject = getChildTasksByParent(task, taskParent)
                else:
                    tasksMasterProject = getChildTasks(task)
                    
                print(f"Task: {node.strip()}")
            else:
                print(f"Task: {node.strip()} not found")
                tasksMasterProject = []
                break
        else:
            print(f"Task: {node.strip()} not found")
            tasksMasterProject = []
            break
        indexBaseNode += 1
        
    print('--------------------------------------')
    mpp.FileSave()
    return tasksMasterProject, taskParent

def updatePowerPoint(sourcePptx, indexSlides, tasksMSProject, taskParentMSProject):
    try:  
        pptx: w3c.CDispatch = w3c.Dispatch("PowerPoint.Application")
        try:
            presentation = pptx.Presentations.Open(sourcePptx.path)
            
            for indexSlide in indexSlides: 
                print() 
                print() 
                print('-------------------------')
                print(f'Update slide: {indexSlide}')           
                
                indexSlide = int(indexSlide)     
                slide = presentation.Slides.Item(indexSlide)  
                if slide is not None:                                  
                    # Walk through all shapes
                    for shape in slide.Shapes:
                        # Is table
                        if shape is not None and shape.HasTable and shape.Table is not None: 
                            table = shape.Table
                            numRows = table.Rows.Count
                            numColumns = table.Columns.Count
                            
                            print('-------------------------')    
                            print(f"Table: {numRows} rows x {numColumns} columns")
                            
                            # Headers
                            headers = []
                            for column in range(1, numColumns + 1):
                                contentCell = table.Cell(1, column).Shape.TextFrame.TextRange.Text
                                print(f'Header: {contentCell}')
                                headers.append({
                                    'column': column,
                                    'content': contentCell
                                })
                                                    
                            headerId = next((x for x in headers
                                        if x is not None and x['content'] == 'Id'), None)
                            
                            headerActivity = next((x for x in headers
                                        if x is not None and x['content'] == 'Actividad'), None)
                            
                            headerDesviation =  next((x for x in headers 
                                        if x is not None and x['content'] == '%Desvío'), None)
                            
                            headerReal =  next((x for x in headers
                                        if x is not None and x['content'] == '%Real'), None)   
                            
                            headerDelay =  next((x for x in headers
                                        if x is not None and x['content'] == 'Motivo del atraso'), None)  
                            
                            headerActionExecuted =  next((x for x in headers
                                        if x is not None and x['content'] == 'Acción a ejecutar'), None)
                            
                            if headerId is not None and headerActivity is not None and headerDesviation is not None and headerReal is not None and headerDelay is not None and headerActionExecuted is not None:
                                                                           
                                for row in range(2, numRows + 1):
                                    id = table.Cell(row, headerId['column']).Shape.TextFrame.TextRange.Text
                                    activity = table.Cell(row, headerActivity['column']).Shape.TextFrame.TextRange.Text
                                    idActivity = f'{activity}'
                                    
                                    taskMSProject = next((x for x in tasksMSProject
                                                            if x is not None 
                                                            and idActivity.strip().upper() in x.Name.upper()
                                                            and getParentTask(x) == taskParentMSProject.Name),
                                                        None)

                                    if taskMSProject is not None:
                                        percentageReal = taskMSProject.PercentComplete
                                        percentageDeviation = taskMSProject.Text1
                                        listNotes = taskMSProject.Notes.split('#')
                                        delay = listNotes[0].strip() if len(listNotes) > 0 else ''
                                        actionExecuted = listNotes[1].strip() if len(listNotes) > 1 else ''
                                        
                                        table.Cell(row, headerReal['column']).Shape.TextFrame.TextRange.Text = f'{percentageReal}%'
                                        table.Cell(row, headerDesviation['column']).Shape.TextFrame.TextRange.Text = f'{percentageDeviation}'
                                        table.Cell(row, headerDelay['column']).Shape.TextFrame.TextRange.Text = f'{delay}'
                                        table.Cell(row, headerActionExecuted['column']).Shape.TextFrame.TextRange.Text = f'{actionExecuted}'
                                        
                                        # add circle 
                                        cell =   table.Cell(row, headerDesviation['column']).Shape
                                        height = (cell.Height - 11) / 2
                                        
                                        circle = slide.Shapes.AddShape(9, (cell.left + 10), (cell.top + height), 11, 11)
                                        # int percentageDeviation
                                        if '%' in percentageDeviation:
                                            percentageDeviation = percentageDeviation.replace('%', '')
                                        
                                        percentageDeviation = int(percentageDeviation)
                                          
                                        # validation color circle
                                        if percentageDeviation <= 3:
                                            circle.Fill.ForeColor.RGB = 49160
                                        
                                        if percentageDeviation > 3 and percentageDeviation <= 6:
                                            circle.Fill.ForeColor.RGB = 65535
                                            
                                        if percentageDeviation > 6:
                                            circle.Fill.ForeColor.RGB = 255
                                    
                                        print(f"Cell update: {idActivity} ({percentageReal}% real) ({percentageDeviation}% deviation)")
         
                                    else:
                                        print(f"Cell error: {idActivity} (Not found in MS Project)")
                                
                            else:
                                print('Headers not found')
                                                                                                                                                        
                else:
                    print(f'Slide {indexSlide} not found')
               
            presentation.SaveAs(sourcePptx.path)
            print()
            print(f'Save PowerPoint')
            presentation.Close()
            
        except Exception as e:
            print(str(e))
                
    except Exception as  e:
        print(str(e))
        
    finally:
        pptx.Quit()