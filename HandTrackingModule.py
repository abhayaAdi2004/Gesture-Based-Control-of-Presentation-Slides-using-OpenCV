"""
HandTrackingModule.py
----------------------
Hand detection using MediaPipe and FingerUp logic.
Detects 21 hand landmarks and determines which fingers are up
using landmark position comparisons (no ML classifier).

FingerUp Logic (replaces CNN):
    Thumb:  tip(landmark 4).x  >  joint(landmark 3).x  => Open
    Index:  tip(landmark 8).y  <  joint(landmark 6).y  => Open
    Middle: tip(landmark 12).y <  joint(landmark 10).y => Open
    Ring:   tip(landmark 16).y <  joint(landmark 14).y => Open
    Pinky:  tip(landmark 20).y <  joint(landmark 18).y => Open
"""

import cv2
import mediapipe as mp


class HandDetector:
    """
    Detects hands using MediaPipe and determines finger states
    using landmark coordinate comparisons.
    """

    def __init__(self, detection_confidence=0.5):
        self.mp_hands = mp.solutions.hands
        self.hands = self.mp_hands.Hands(
            static_image_mode=False,
            max_num_hands=1,
            min_detection_confidence=detection_confidence,
        )
        self.mp_draw = mp.solutions.drawing_utils
        self.landmarks = []
        # Landmark indices: 4=ThumbTip, 8=IndexTip, 12=MiddleTip,
        #                   16=RingTip, 20=PinkyTip

    def find_hands(self, img, draw=True):
        """Detect hands and optionally draw landmarks."""
        img = cv2.flip(img, 1)  # Mirror view
        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        self.results = self.hands.process(img_rgb)

        if self.results.multi_hand_landmarks:
            for hand_landmarks in self.results.multi_hand_landmarks:
                if draw:
                    self.mp_draw.draw_landmarks(
                        img, hand_landmarks, self.mp_hands.HAND_CONNECTIONS
                    )
        return img

    def find_landmarks(self, img, hand_no=0):
        """Extract 21 landmark positions as [id, x, y]."""
        self.landmarks = []
        h, w, _ = img.shape

        if self.results.multi_hand_landmarks:
            hand = self.results.multi_hand_landmarks[hand_no]
            for idx, lm in enumerate(hand.landmark):
                self.landmarks.append([idx, int(lm.x * w), int(lm.y * h)])
        return self.landmarks

    def fingers_up(self):
        """
        FingerUp logic: Compare landmark positions to determine
        which fingers are raised. Returns [thumb, index, middle, ring, pinky].

        Thumb: tip(4).x > joint(3).x
        Others: tip.y < pip_joint.y (tip is above the joint)
        """
        fingers = [0, 0, 0, 0, 0]

        if len(self.landmarks) < 21:
            return fingers

        # Thumb (landmark 4 vs landmark 3)
        if self.landmarks[4][1] > self.landmarks[3][1]:
            fingers[0] = 1

        # Index (landmark 8 vs landmark 6)
        if self.landmarks[8][2] < self.landmarks[6][2]:
            fingers[1] = 1

        # Middle (landmark 12 vs landmark 10)
        if self.landmarks[12][2] < self.landmarks[10][2]:
            fingers[2] = 1

        # Ring (landmark 16 vs landmark 14)
        if self.landmarks[16][2] < self.landmarks[14][2]:
            fingers[3] = 1

        # Pinky (landmark 20 vs landmark 18)
        if self.landmarks[20][2] < self.landmarks[18][2]:
            fingers[4] = 1

        return fingers

