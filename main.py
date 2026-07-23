"""
main.py - Gesture Controlled PowerPoint Presentation
====================================================

Workflow (as described in the paper):
    Start Program
        -> Open Webcam
        -> Capture Frame
        -> Detect Hand
        -> Detect 21 Landmarks
        -> FingerUp Function
        -> Recognize Gesture
        -> Gesture Mapping
        -> PyWin32 (PowerPoint API)
        -> Next Slide / Previous Slide
        -> Repeat

Technologies: OpenCV, MediaPipe Hands, PyWin32
"""

import cv2
import sys
import time
from HandTrackingModule import HandDetector

try:
    import win32com.client
except ImportError:
    print("[ERROR] pywin32 is required. Install: pip install pywin32")
    sys.exit(1)


class PPTController:
    """
    Controls PowerPoint using the COM API (PyWin32).
    Uses Presentation.SlideShowWindow.View.Next() / Previous()
    as described in the paper.
    """

    def __init__(self):
        self.app = None
        self.presentation = None
        self.slide_show = None
        self.view = None
        self.active = False
        self._connect()

    def _connect(self):
        """Connect to a running PowerPoint instance."""
        try:
            self.app = win32com.client.Dispatch("PowerPoint.Application")
            self.app.Visible = True
            print("[INFO] Connected to PowerPoint.")
        except Exception as e:
            print(f"[ERROR] Could not connect to PowerPoint: {e}")

    def start(self):
        """Start the slideshow (F5)."""
        try:
            self.app.Activate()
            # Send F5 key
            import win32api
            import win32con
            win32api.keybd_event(win32con.VK_F5, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(win32con.VK_F5, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.5)
            self._get_view()
            self.active = True
            print("[ACTION] Slideshow started.")
        except Exception as e:
            print(f"[WARNING] start() failed: {e}")

    def _get_view(self):
        """Get the SlideShowWindow.View object for Next()/Previous()."""
        try:
            if self.app and self.app.SlideShowWindows.Count > 0:
                self.slide_show = self.app.SlideShowWindows(1)
                self.view = self.slide_show.View
                self.active = True
            else:
                self.view = None
                self.active = False
        except Exception:
            self.view = None
            self.active = False

    def next_slide(self):
        """Next slide using View.Next() - the paper's PyWin32 method."""
        if not self.active:
            self._get_view()
        if self.view:
            try:
                self.view.Next()
                print("[ACTION] Next slide")
            except Exception as e:
                print(f"[WARNING] Next() failed: {e}")

    def previous_slide(self):
        """Previous slide using View.Previous() - the paper's PyWin32 method."""
        if not self.active:
            self._get_view()
        if self.view:
            try:
                self.view.Previous()
                print("[ACTION] Previous slide")
            except Exception as e:
                print(f"[WARNING] Previous() failed: {e}")

    def stop(self):
        """Stop slideshow (Escape)."""
        if self.active:
            try:
                import win32api
                import win32con
                win32api.keybd_event(win32con.VK_ESCAPE, 0, 0, 0)
                time.sleep(0.1)
                win32api.keybd_event(win32con.VK_ESCAPE, 0, win32con.KEYEVENTF_KEYUP, 0)
                self.active = False
                print("[ACTION] Slideshow stopped.")
            except Exception as e:
                print(f"[WARNING] stop() failed: {e}")

    def close(self):
        self.app = None
        print("[INFO] PPT controller closed.")


def main():
    print("=" * 50)
    print("  Gesture Controlled PowerPoint Presentation")
    print("  Based on: FingerUp Logic + PyWin32 API")
    print("=" * 50)

    # Camera
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("[ERROR] No webcam found.")
        sys.exit(1)

    # Hand detector
    detector = HandDetector(detection_confidence=0.7)

    # PowerPoint controller
    ppt = PPTController()

    # ---- Step 8: Prevent Multiple Slide Changes ----
    # Uses a frame-based delay counter (20 frames wait)
    # to prevent one gesture from changing 30 slides.
    delay_counter = 0
    DELAY_FRAMES = 20

    print("\n[READY] Show Index finger [0,1,0,0,0] for Next slide.")
    print("       Show Thumb [1,0,0,0,0] for Previous slide.")
    print("       Show Open Palm to start slideshow.")
    print("       Press 'q' in the camera window to quit.\n")

    # ---- Step 10: Continuous Loop ----
    while True:
        success, img = cap.read()
        if not success:
            break

        # Capture frame -> Detect hand -> Extract 21 landmarks
        img = detector.find_hands(img, draw=True)
        landmarks = detector.find_landmarks(img)

        gesture_text = "NONE"
        action_text = ""

        if landmarks:
            # FingerUp Function
            fingers = detector.fingers_up()
            finger_str = f"[{fingers[0]},{fingers[1]},{fingers[2]},{fingers[3]},{fingers[4]}]"

            # Recognize Gesture
            if fingers == [1, 0, 0, 0, 0]:
                gesture_text = "THUMBS_UP"
            elif fingers == [0, 1, 0, 0, 0]:
                gesture_text = "POINT"
            elif fingers == [1, 1, 1, 1, 1]:
                gesture_text = "OPEN_PALM"
            else:
                gesture_text = "NONE"

            # Gesture Mapping + Delay Counter
            if delay_counter > 0:
                delay_counter -= 1
                action_text = f"waiting ({delay_counter})"

            elif gesture_text == "POINT":
                # [0,1,0,0,0] -> Next Slide
                ppt.next_slide()
                delay_counter = DELAY_FRAMES
                action_text = "NEXT"

            elif gesture_text == "THUMBS_UP":
                # [1,0,0,0,0] -> Previous Slide
                ppt.previous_slide()
                delay_counter = DELAY_FRAMES
                action_text = "PREV"

            elif gesture_text == "OPEN_PALM":
                ppt.start()
                delay_counter = DELAY_FRAMES
                action_text = "START"

            # Draw finger state
            cv2.putText(img, f"Fingers: {finger_str}", (10, 100),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 200, 0), 2)
            cv2.putText(img, f"Gesture: {gesture_text}", (10, 130),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)
            if action_text:
                cv2.putText(img, f"Action: {action_text}", (10, 160),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 255), 2)
        else:
            # No hand detected
            cv2.putText(img, "No hand detected", (10, 100),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)

        # Instructions
        cv2.putText(img, "Index=Next  Thumb=Prev  Palm=Start  q=Quit",
                    (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (200, 200, 200), 2)

        cv2.imshow("Gesture PPT Control", img)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    ppt.close()
    print("[INFO] Program ended.")


if __name__ == "__main__":
    main()

