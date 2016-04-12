import cv2
import numpy as np
import operator
import os

# module level variables ##########################################################################
MIN_CONTOUR_AREA = 50
RESIZED_IMAGE_WIDTH = 20
RESIZED_IMAGE_HEIGHT = 20
PATH = "Captcha.jpg"

class ContourWithData():

    npaContour = None           # contour
    boundingRect = None         # bounding rect for contour
    intRectX = 0                # bounding rect top left corner x location
    intRectY = 0                # bounding rect top left corner y location
    intRectWidth = 0            # bounding rect width
    intRectHeight = 0              # bounding rect height
    fltArea = 0.0               # area of contour

    def calculateRectTopLeftPointAndWidthAndHeight(self):               # calculate bounding rect info
        [intX, intY, intWidth, intHeight] = self.boundingRect
        self.intRectX = intX
        self.intRectY = intY
        self.intRectWidth = intWidth
        self.intRectHeight = intHeight

    def checkIfContourIsValid(self):
        if self.fltArea < MIN_CONTOUR_AREA: return False        # much better validity checking would be necessary
        return True

def read_captcha(Captcha=None):
    allContoursWithData = []
    validContoursWithData = []
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    npaFlattenedImages= np.loadtxt(os.path.join(BASE_DIR, "ocr/npFlattenedCombined.txt"),np.float32) 
    npaClassifications= np.loadtxt(os.path.join(BASE_DIR, "ocr/npClassificationCombined.txt"), np.float32)
    kNearest = cv2.KNearest()
    kNearest.train(npaFlattenedImages, npaClassifications)
    #Captcha = cv2.imread(PATH)

    imgGray = cv2.cvtColor(Captcha, cv2.COLOR_BGR2GRAY)
    imgBlurred = cv2.GaussianBlur(imgGray, (5,5), 0)
    imgThresh = cv2.adaptiveThreshold(imgBlurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV,11,2)
    imgThreshCopy = imgThresh.copy()



    npaContours, npaHierarchy = cv2.findContours(imgThreshCopy, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for npaContour in npaContours:
        contourWithData = ContourWithData()                                             # instantiate a contour with data object
        contourWithData.npaContour = npaContour                                         # assign contour to contour with data
        contourWithData.boundingRect = cv2.boundingRect(contourWithData.npaContour)     # get the bounding rect
        contourWithData.calculateRectTopLeftPointAndWidthAndHeight()                    # get bounding rect info
        contourWithData.fltArea = cv2.contourArea(contourWithData.npaContour)           # calculate the contour area
        allContoursWithData.append(contourWithData)

    for contourWithData in allContoursWithData:
        if contourWithData.checkIfContourIsValid():
            validContoursWithData.append(contourWithData)

    validContoursWithData.sort(key = operator.attrgetter("intRectX"))
    strFinalString = ""
    i = 0
    for contourWithData in validContoursWithData:
        i += 1
        cv2.rectangle (Captcha, (contourWithData.intRectX, contourWithData.intRectY),(contourWithData.intRectX + contourWithData.intRectWidth, contourWithData.intRectY + contourWithData.intRectHeight),(0, 255, 0),2)
        imgROI = imgThresh[contourWithData.intRectY : contourWithData.intRectY + contourWithData.intRectHeight,contourWithData.intRectX : contourWithData.intRectX + contourWithData.intRectWidth]
        imgROIResized = cv2.resize(imgROI, (RESIZED_IMAGE_WIDTH, RESIZED_IMAGE_HEIGHT))
        npaROIResized = imgROIResized.reshape((1, RESIZED_IMAGE_WIDTH * RESIZED_IMAGE_HEIGHT))
        npaROIResized = imgROIResized.reshape((1, RESIZED_IMAGE_WIDTH * RESIZED_IMAGE_HEIGHT))
        npaROIResized = np.float32 (npaROIResized)


        retval, npaResults, neigh_resp, dists = kNearest.find_nearest(npaROIResized, k = 1)

        retval = int(retval)	
        
        if retval > 9 :
        	retval = chr(int(retval))
        strCurrentChar = str(retval)
        # print strCurrentChar
        strFinalString = strFinalString + strCurrentChar

        # cv2.namedWindow('Fuck '+str(i),cv2.WINDOW_NORMAL)
        # cv2.imshow('Fuck '+str(i),imgROI)
        # cv2.waitKey(0)

    print strFinalString


    # cv2.imshow("imgTestingNumbers", Captcha)
    # cv2.waitKey(0)
    cv2.destroyAllWindows()
    return strFinalString
            