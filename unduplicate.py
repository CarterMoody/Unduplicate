# Find Windows File properties/metadata
#       Useful for determining which of the duplicate files to keep and which one to throw out
#       https://learn.microsoft.com/en-us/windows/win32/properties/props-system-video-totalbitrate
#       Replace period '.' with underscore '_'
#       https://stackoverflow.com/questions/31507038/python-how-to-read-windows-media-created-date-not-file-creation-date
# 
#
# The findSimilarMatches may need to be tuned based on your specific needs.
#   It is ultimately just a guess, and it's results should be checked manually
#   For example, my files all start with a random word (DVD29920 or S1945E15), 
#   so that must be ripped out before comparing title to determine if they are the same media




import os
import argparse
from difflib import SequenceMatcher
from win32com.propsys import propsys, pscon # > python -m pip install pywin32



# Compares two strings and gives similarity score
def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

# Gets list of files in source directory with specific extensions
# Returns list of matches
def fileList(source):
    matches = []
    for root, dirnames, filenames in os.walk(source):
        for filename in filenames:
            if filename.endswith(('.mov', '.MOV', '.avi', '.mpg', '.mkv', '.mp4', '.wmv', '.flv', '.MP4')):
                matches.append(os.path.join(root, filename))
                #matches.append(filename)    # for some reason os.remove is not working with full path...
    return matches
    

# Goes through fileList and returns a list of pairs that are similar enough to probably be the same media
def findSimilarMatches(fileList, similarityThreshold):
    similarMatches = []
    for index in range(len(fileList)): 
        fileA = fileList[index]
        for index2 in range(index, len(fileList)): # Start comparing from index->onward so we don't double-compare to previous
            fileB = fileList[index2]
            #print(f"fileA: {fileA} fileB: {fileB}")
            fileANoExtension = os.path.splitext(fileA)[0]
            fileAExtension = os.path.splitext(fileA)[1]
            fileBNoExtension = os.path.splitext(fileB)[0]
            fileBExtension = os.path.splitext(fileB)[1]
            fileABaseName = os.path.basename(fileANoExtension)
            fileBBaseName = os.path.basename(fileBNoExtension)
            first, space, rest = fileABaseName.partition(' ') # Rip out first word of the filename and THEN compare
            fileATrueMediaName = rest or first
            #print(f"fileATrueMediaName: {fileATrueMediaName}")
            first, space, rest = fileBBaseName.partition(' ') # Rip out first word of the filename and THEN compare
            fileBTrueMediaName = rest or first
            #print(f"fileBTrueMediaName: {fileBTrueMediaName}")
            similarityScore = similar(fileATrueMediaName, fileBTrueMediaName)

            
            if similarityScore == 1.0:
                # Double check now with extension back on filename
                if fileAExtension != fileBExtension:
                    matchPair = [fileA, fileB]
                    similarMatches.append(matchPair)
                else:
                    continue # We don't want to consider a match on itself
            elif similarityScore >= similarityThreshold:
                print(f"   fileA: {fileATrueMediaName}")
                print(f"   compared to:")
                print(f"   fileB: {fileBTrueMediaName}")
                print(f"   score: {similarityScore}\n")
                # Add logic here to remove either fileA or fileB based on file details/size/bitrate etc..
                matchPair = [fileA, fileB]
                similarMatches.append(matchPair)
        
    return similarMatches
    
   
# Deletes a file
def deleteFile(file):
    os.remove(file)
    print(f"deleted {file}")


# iterates through all similarMatches and deletes the worse file
def inspectMediaPairs(similarMatches):
    for mediaPair in similarMatches:
        # Check if both files exist in the mediaPair
        if os.path.exists(mediaPair[0]) and os.path.exists(mediaPair[1]): # Should this check be in the chooseBetterFile function instead?
            worseFile = chooseBetterFile(mediaPair)
            deleteFile(worseFile)
        # Need to remove any other pairs that include the worseFile that was just deleted?
        
    
# Goes through list of pairs of similar matches and returns the one to be deleted
def chooseBetterFile(mediaPair):
    fileA = mediaPair[0]
    fileB = mediaPair[1]
    lowerBitrateFile = compareBitrate(fileA, fileB)
    return lowerBitrateFile

 
# Accepts two files and compares their bitrates, returning the worse file
def compareBitrate(fileA, fileB):
        propertiesA = propsys.SHGetPropertyStoreFromParsingName(fileA)
        bitrateEncodingA = propertiesA.GetValue(pscon.PKEY_Video_EncodingBitrate).GetValue()
        #bitrateTotalA = propertiesA.GetValue(pscon.PKEY_Video_TotalBitrate).GetValue()
        propertiesB = propsys.SHGetPropertyStoreFromParsingName(fileB)
        #bitrateTotalB = propertiesB.GetValue(pscon.PKEY_Video_EncodingBitrate).GetValue()
        bitrateEncodingB = propertiesB.GetValue(pscon.PKEY_Video_EncodingBitrate).GetValue()
        
        print(f"choosing between: ")
        print(f"   fileA: {fileA}")
        print(f"      bitrate: {bitrateEncodingA}")
        print(f"   fileB: {fileB}")
        print(f"      bitrate: {bitrateEncodingB}")
        
        if bitrateEncodingA < bitrateEncodingB:
            print(f"   fileA has lower bitrate, delete it!")
            return fileA
   
        elif bitrateEncodingB < bitrateEncodingA:
            print(f"   fileB has lower bitrate, delete it!")
            return fileB
 
        elif bitrateEncodingA == bitrateEncodingB:
            print(f"   fileA has the same bitrate as fileB!")
            fileABaseName = os.path.basename(fileA).lower()
            fileBBaseName = os.path.basename(fileB).lower()
            if fileABaseName.startswith("s"):   # If the name specifies season information, that's likely worth keeping
                return fileB
            else:
                return fileA

 
def main():  
    cwd = os.getcwd()
    if args["directory"]:
        print(f"cleaning directory {args['directory']}")
        directory = args["directory"]
    else:
        print("no user supplied directory, using CWD recursively!")
        directory = cwd   
    fileMatches = fileList(directory)
    print(f"fileMatches: {fileMatches}")

    similarityThreshold = 0.85
    similarMatches = findSimilarMatches(fileMatches, similarityThreshold)
    print(f"similarMatches: {similarMatches}")

    inspectMediaPairs(similarMatches)


# Construct the argument parser and parse the arguments
arg_desc = '''\
        Let's load an image from the command line!
        --------------------------------
            This program loads an image
            with OpenCV and Python argparse!
        '''
parser = argparse.ArgumentParser(formatter_class = argparse.RawDescriptionHelpFormatter,
                                    description= arg_desc)
 
parser.add_argument("-d", "--directory", metavar="PATH", help = "FULL Path to your directory to be cleaned like 'C:\scripts\test'")
args = vars(parser.parse_args())


if __name__=="__main__":
    main()