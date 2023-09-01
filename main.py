import argparse
import os
import time
from functions.pptx import updatePowerPoint, fileExists, getTasksByName

def main(args):
    timeStart = time.time()
    
    sourceFolder: str = args.source_folder
    indexSlides: list = args.index_slides.split(',')
    baseNode: str = args.base_node
    
    sourcePptx, sourceMpp = fileExists(sourceFolder)
    print(f'Source file PowerPoint: {sourcePptx.name}')
    print(f'Source file MS Project: {sourceMpp.name}')

    tasksMSProject, taskParentMSProject = getTasksByName(baseNode, sourceMpp)
    
    updatePowerPoint(sourcePptx, indexSlides, tasksMSProject, taskParentMSProject)

    timeEnd = time.time()
    minutes = (timeEnd - timeStart) / 60
    print('--------------------------------------')
    print(f'Time: {minutes} minutes')

if __name__ == '__main__':
    parser: argparse.ArgumentParser = argparse.ArgumentParser()
    # Argument source folder
    parser.add_argument('-sf',
                        '--source_folder',
                        required=True,
                        help='Source folder with file Power Point and MS Project',
                        type=str)
    
    # Argument index slides
    parser.add_argument('-is',
                        '--index_slides',
                        required=True,
                        help='Index slides to update',
                        type=str) 
    
    # Argument base node MS Project
    parser.add_argument('-bn',
                        '--base_node',
                        required=True,
                        help='Base node MS Project',
                        type=str)

    args: argparse.Namespace = parser.parse_args()
    
    if args.source_folder and os.path.exists(args.source_folder):
        if args.index_slides and len(args.index_slides) > 0:
            main(args)
        else:
            print('Error: The index slides is empty')
    else:
        print('Error: The source folder does not exist')

    