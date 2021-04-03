import cv2
import sys
import os
from PIL import Image

refPt = []
cropping = False
def click_and_crop(event, x, y, flags, param):
	
	global refPt, cropping
	
	if event == cv2.EVENT_LBUTTONDOWN:
		refPt = [(x, y)]
		cropping = True
	
	elif event == cv2.EVENT_LBUTTONUP:
		refPt.append((x, y))
		poi = open("co-ordinates.txt","w")
		left,top = refPt[0]
		right,bottom = refPt[1]
		poi.write(str(refPt[0][1])+','+str(refPt[1][1])+','+str(refPt[0][0])+','+str(refPt[1][0]))
		poi.close()
		cropping = False
		cv2.rectangle(image, refPt[0], refPt[1], (0, 255, 0), 3)
		cv2.imshow("image", image)


image = cv2.imread(sys.argv[1])
clone = image.copy()
cv2.namedWindow("image", cv2.WINDOW_NORMAL)
cv2.setMouseCallback("image", click_and_crop)

while True:
	
	cv2.imshow("image", image)
	key = cv2.waitKey(1) & 0xFF
	if key == ord("r"):
		image = clone.copy()
	elif key == ord("c"):
		break

if len(refPt) == 2:
	roi = clone[refPt[0][1]:refPt[1][1],refPt[0][0]:refPt[1][0]]
	cv2.imshow("ROI", roi)
	cv2.imwrite(os.getcwd()+"\\cropped\\{}.jpg".format(sys.argv[2]), roi)
	cv2.waitKey(0)

cv2.destroyAllWindows()