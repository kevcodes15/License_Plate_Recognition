import cv2
import easyocr
import pandas as pd

harcascade = "model\haarcascade_russian_plate_number.xml"

cap = cv2.VideoCapture(0)

cap.set(3, 640) # width
cap.set(4, 480) #height

min_area = 500
count = 0

while True:
    success, img = cap.read()

    plate_cascade = cv2.CascadeClassifier(harcascade)
    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    plates = plate_cascade.detectMultiScale(img_gray, 1.1, 4)

    for (x,y,w,h) in plates:
        area = w * h

        if area > min_area:
            cv2.rectangle(img, (x,y), (x+w, y+h), (0,255,0), 2)
            cv2.putText(img, "Number Plate", (x,y-5), cv2.FONT_HERSHEY_COMPLEX_SMALL, 1, (255, 0, 255), 2)

            img_roi = img[y: y+h, x:x+w]
            cv2.imshow("ROI", img_roi)


    reader = easyocr.Reader(['en'])
    cv2.imshow("Result", img)

    if cv2.waitKey(1) & 0xFF == ord('s'):
        cv2.imwrite("plates/scanned_img_" + str(count) + ".jpg", img_roi)

        output = reader.readtext("plates/scanned_img_" + str(count) + ".jpg")
        plate_number = output[0][1][1:]  # Extracting just the number plate value

        # Existing Excel file
        existing_file = 'database.xlsx'
 
        # New data to append
        new_data = {'Car No.': [str('Car_no_' + str(count))], 'Plate Number': [str(plate_number)]}
        df_new = pd.DataFrame(new_data)
 
        # Read existing data
        df_existing = pd.read_excel(existing_file)
 
        # Append new data
        df_combined = df_existing.append(df_new, ignore_index=True)
 
        # Save the combined data to Excel
        df_combined.to_excel(existing_file, index=False)       


        cv2.rectangle(img, (0,200), (640,300), (0,255,0), cv2.FILLED)
        cv2.putText(img, "Plate Saved", (150, 265), cv2.FONT_HERSHEY_COMPLEX_SMALL, 2, (0, 0, 255), 2)
        cv2.imshow("Results",img)
        cv2.waitKey(500)
        count += 1