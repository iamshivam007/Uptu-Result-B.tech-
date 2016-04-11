import numpy as np
import cv2
RESIZED_IMAGE_WIDTH = 20
RESIZED_IMAGE_HEIGHT = 20
PATH = "Char5.jpg"

def preprocessing(img):
	imgGray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
	imgBlurred = cv2.medianBlur(imgGray, 1, 0)
	imgThresh = cv2.adaptiveThreshold(imgBlurred,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY_INV,11,2)
	imgThresh = cv2.resize (imgThresh, (RESIZED_IMAGE_WIDTH, RESIZED_IMAGE_HEIGHT))
	return imgThresh

npFlattenedImages = np.loadtxt ("npFlattenedCombined.txt", np.float32)
npClassifications = np.loadtxt("npClassificationCombined.txt", np.float32)
npClassifications = npClassifications.reshape((npClassifications.size, 1))
kNearest = cv2.KNearest()
kNearest.train (npFlattenedImages, npClassifications)
img = cv2.imread(PATH)
imgThresh = preprocessing(img)

# print imgThresh.shape
npROIResized = np.float32(imgThresh.reshape((1, RESIZED_IMAGE_WIDTH * RESIZED_IMAGE_HEIGHT )))
# print npaROIResized.shape
retval, npResults, neigh_resp, dists = kNearest.find_nearest(npROIResized, k = 1)

if (retval >= 65):
	print chr(int(retval))

else:
	print retval