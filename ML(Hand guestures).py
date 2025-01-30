import cv2
import mediapipe as mp
import pyautogui
import time

# Initialize MediaPipe Hands.
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(max_num_hands=2, min_detection_confidence=0.7)
mp_draw = mp.solutions.drawing_utils

# Initialize the webcam
cap = cv2.VideoCapture(0)

# Open PowerPoint and wait for it to load
pyautogui.press('win')
time.sleep(1)  # Wait for the start menu to open
pyautogui.write('PowerPoint')  # Search for PowerPoint
pyautogui.press('enter')
time.sleep(5)  # Wait for PowerPoint to open

# Control PowerPoint presentation
pyautogui.press('f5')  # Start slideshow
time.sleep(2)  # Wait for the slideshow to start

while True:
    success, image = cap.read()
    if not success:
        print("Ignoring empty camera frame.")
        continue

    # Convert the image color from BGR to RGB.
    image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    image.flags.writeable = False

    # Process the image and detect hands
    results = hands.process(image)

    # Draw the hand annotations on the image.
    image.flags.writeable = True
    image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            mp_draw.draw_landmarks(
                image, hand_landmarks, mp_hands.HAND_CONNECTIONS)

            # Control PowerPoint slides based on hand gestures
            if hand_landmarks.landmark:
                # Get the y-coordinate of the middle finger tip
                middle_finger_tip_y = hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_TIP].y

                # If the middle finger tip is above a certain threshold, go to the next slide
                if middle_finger_tip_y < 0.5:
                    pyautogui.press('right')  # Go to the next slide
                    time.sleep(1)  # Add a small delay to prevent multiple key presses

                # If the middle finger tip is below a certain threshold, go to the previous slide
                elif middle_finger_tip_y > 0.5:
                    pyautogui.press('left')  # Go to the previous slide
                    time.sleep(1)  # Add a small delay to prevent multiple key presses

    # Display the resulting frame
    cv2.imshow('Hand Tracking', image)
    if cv2.waitKey(5) & 0xFF == 27:
        break

# Release the webcam and destroy all OpenCV windows.
cap.release()
cv2.destroyAllWindows()

# Exit PowerPoint slideshow
pyautogui.press('esc')  # Exit slideshow

# Close PowerPoint
pyautogui.hotkey('alt', 'f4')  # Close PowerPoint
