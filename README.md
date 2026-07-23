# Gesture Controlled PowerPoint Presentation

Control PowerPoint slides using hand gestures via webcam.  
Implements the FingerUp mathematical logic (replaces a trained ML classifier).

## Architecture — Complete Workflow

```
Start Program
    │
    ▼
Open Webcam (OpenCV)
    │
    ▼
Capture Frame
    │
    ▼
Detect Hand (MediaPipe Hands)
    │
    ▼
Detect 21 Landmarks
    │
    ▼
FingerUp Function (landmark comparison)
    │
    ▼
Recognize Gesture
    │
    ▼
Gesture Mapping
    │
    ▼
PyWin32 (PowerPoint COM API)
    │
    ▼
Next Slide / Previous Slide
    │
    ▼
Repeat (Continuous Loop)
```

## Core Technologies

- **OpenCV** — Camera capture and frame processing
- **MediaPipe Hands** — 21 hand landmark detection
- **PyWin32** — PowerPoint COM API (`SlideShowWindow.View.Next() / Previous()`)

## FingerUp Logic (Replaces CNN/ML Classifier)

| Finger | Tip Landmark | Joint Landmark | Open Condition |
|--------|-------------|----------------|----------------|
| Thumb | 4 | 3 | `tip.x > joint.x` |
| Index | 8 | 6 | `tip.y < joint.y` |
| Middle | 12 | 10 | `tip.y < joint.y` |
| Ring | 16 | 14 | `tip.y < joint.y` |
| Little | 20 | 18 | `tip.y < joint.y` |

Output is a binary array `[thumb, index, middle, ring, pinky]`

## Gestures Used

| Gesture | Finger Pattern | Action |
|---------|---------------|--------|
| Index Finger Up | `[0,1,0,0,0]` | **Next Slide** |
| Thumb Up | `[1,0,0,0,0]` | **Previous Slide** |
| Open Palm | `[1,1,1,1,1]` | Start Slideshow |

## Accuracy (Reported in Paper)

| Distance | Next Slide | Previous Slide |
|----------|-----------|----------------|
| 15 cm | 99% | 98% |
| 25 cm | 98% | 97% |
| 45 cm | 97% | 96% |
| 60 cm | 97% | 93% |

**Average accuracy: ~97%**

**Known limitation:** Recognizing the back of the hand.

## Slide Change Protection

A 20-frame delay counter prevents one gesture from triggering multiple slide changes.

## Prerequisites

- Windows OS (for PyWin32)
- Python 3.10 or 3.11
- Webcam
- Microsoft PowerPoint installed

## Installation

```bash
cd "Gesture PPT Control"
pip install -r requirements.txt
```

## Usage

```bash
python main.py
```

Controls:
- **Index finger up** `[0,1,0,0,0]` → Next slide
- **Thumb up** `[1,0,0,0,0]` → Previous slide
- **Open palm** `[1,1,1,1,1]` → Start slideshow
- Press **'q'** to quit

## Project Structure

```
Gesture PPT Control/
├── main.py                  # Main loop: capture → detect → FingerUp → PPT command
├── HandTrackingModule.py    # MediaPipe hand detection + FingerUp logic
├── requirements.txt         # Python dependencies
└── README.md                # This file
```

