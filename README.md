# From Gestures to Words: A Hybrid System for Interactive PowerPoint Presentations

**Author:** Biddut Sarker Bijoy

## Description

A hand gesture and speech-controlled presentation viewer that allows us to navigate slides, to draw anything, use a laser pointer, and transcribe speech in real-time without requiring PowerPoint installation.

## Features

- **Hand Gesture Control** - Navigate and interact with slides using hand gestures
- **Real-time Drawing** - Draw annotations on slides with multiple colors
- **Laser Pointer** - Highlight content with a virtual laser pointer
- **Eraser Tool** - Remove drawings with gesture control
- **Speech-to-Text** - Real-time speech transcription during presentations
- **PPTX Support** - Direct PPTX file rendering (no PowerPoint needed)

## Requirements

### Python packages
```bash
pip install opencv-python mediapipe python-pptx pillow pdf2image SpeechRecognition pyaudio
```

### System dependencies (macOS)
```bash
brew install poppler libreoffice
```

## Installation

1. Install system dependencies (Required):
```bash
brew install poppler libreoffice
```

## Usage

Run the application:
```bash
python3 pptx_controller.py
```

Enter the path to your PPTX file when prompted.

## Hand Gestures

| Gesture | Pattern | Action |
|---------|---------|--------|
| Index finger only | `{0,1,0,0,0}` | Draw on slide |
| Index + Middle | `{0,1,1,0,0}` | Laser pointer |
| Index + Middle + Ring | `{0,1,1,1,0}` | Erase |
| Thumb only | `{1,0,0,0,0}` | Previous slide |
| Open palm (all fingers) | `{1,1,1,1,1}` | Next slide |
| Index + Pinky | `{0,1,0,0,1}` | Change color |
| Thumb + Index + Pinky | `{1,1,0,0,1}` | Clear drawings |

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| `Q` | Quit application |
| `N` | Next slide |
| `P` | Previous slide |
| `C` | Clear drawings on current slide |

## Drawing Colors

Red, Green, Blue, Yellow, Cyan, Magenta, White, Orange, Purple, Black

## Windows

- **Presentation Window** - Displays slides with annotations and transcript bar
- **Camera Window** - Shows hand tracking and gesture recognition with instructions

## Code Structure

### Configuration

Global constants defining system behavior:
- `WINDOW_WIDTH`, `WINDOW_HEIGHT` - Display dimensions
- `GESTURE_COOLDOWN` - Prevents gesture double-triggering
- `DRAWING_THICKNESS`, `ERASER_SIZE` - Annotation parameters
- `DRAWING_COLORS` - Color palette dictionary (10 colors in BGR format)

### PPTX Rendering Module

Functions for converting PowerPoint files to image arrays:
- `pptx_to_images_pdf(pptx_path)` - Converts PPTX to PDF using LibreOffice, then converts PDF to images using pdf2image
- `load_pptx(pptx_path)` - Main loader that calls the PDF conversion method

### Gesture Recognition & Distance Estimation Module

Functions for hand tracking and gesture interpretation:
- `estimate_hand_distance(hand_landmarks, frame_width, frame_height)` - Calculates real-time distance from hand to camera in centimeters using hand size and MediaPipe z-coordinates. Returns distance_cm and hand_width_pixels. Optimal range: 30-60cm (green), acceptable: 20-80cm (yellow), otherwise red.
- `count_fingers(hand_landmarks)` - Detects extended fingers using MediaPipe landmarks, returns tuple of 5 binary values
- `detect_gesture(hand_landmarks)` - Maps finger patterns to gestures (draw, laser_pointer, erase, next, previous, clear_all, change_color)

### Speech Recognition Module

Class for real-time speech-to-text transcription:
- **SpeechTranscriber class**
  - `start()` - Initializes background listening thread
  - `stop()` - Terminates speech recognition
  - `_listen_loop()` - Continuous audio capture and recognition
  - `get_transcript()` - Returns latest recognized text
  - `get_full_history()` - Returns last 10 recognized phrases

### Presentation Viewer Module

Main presentation management class:
- **PresentationViewer class**
  - `get_current_color_name()` - Returns the name of the current drawing color
  - `cycle_color()` - Switches to next drawing color from 10 available colors
  - `get_current_frame()` - Renders slide with drawings, laser pointer, eraser cursor, and transcript
  - `next_slide()`, `previous_slide()` - Slide navigation with cooldown
  - `draw_point(x, y)` - Adds drawing on current slide using current color
  - `erase_area(x, y)` - Removes drawing in specified area
  - `clear_all_drawings()` - Clears annotations on current slide
  - `update_laser_pointer(x, y)` - Updates laser position
  - `clear_laser_pointer()` - Clears laser pointer
  - `_add_transcript_bar()` - Renders YouTube-style subtitle overlay at bottom of slide with semi-transparent black background

### UI Overlay Module

Functions for camera feed visual feedback:
- `draw_gesture_info(frame, gesture, slide_info, fingers, current_color_name, distance_cm)` - Displays finger patterns at top, active gesture mode with color coding (green for draw, red for laser, orange for erase, magenta for color change, yellow for others), slide number, current color, distance measurement in top-right with color-coded box (green: optimal 30-60cm, yellow: acceptable 20-80cm, red: too close/far), and comprehensive instruction guide at bottom
- `draw_speech_status(frame)` - Shows microphone status indicator (MIC: ON in green) in top-right corner

### Main Application

`main()` function orchestrates the entire application:
- File path input and validation
- Presentation loading via `load_pptx()`
- Camera initialization (OpenCV VideoCapture at 640x480)
- Speech transcriber setup (if available)
- Creates two windows: 'Presentation' and 'Camera (Hand Tracking)'
- Main event loop:
  - Hand landmark processing via MediaPipe
  - Distance calculation using `estimate_hand_distance()`
  - Gesture detection and action execution with cooldown timers
  - Speech transcript updates
  - Frame rendering for presentation (with laser pointer glow effect, eraser cursor with crosshair, and transcript bar) and camera windows
  - Keyboard input handling (Q=Quit, N=Next, P=Previous, C=Clear)

### Key Data Flows

1. Camera → MediaPipe → Distance Estimation → Gesture Detection → Action
2. Microphone → Speech Recognition → Transcript Queue → Display in Presentation Window
3. PPTX File → PDF Conversion → Image Conversion → Slide Array → Viewer
4. Drawing Actions → Layer Overlay (with color) → Presentation Frame
5. Hand Distance → Real-time Display → User Feedback (Color-coded)

## Technical Details

| Parameter | Value |
|-----------|-------|
| Presentation window resolution | 1280x720 |
| Camera resolution | 640x480 |
| Hand tracking | MediaPipe with single hand detection (min_detection_confidence=0.7, min_tracking_confidence=0.7, model_complexity=1) |
| Distance measurement | Uses hand width in pixels and MediaPipe z-coordinates with focal length of 600px (calibrated for typical webcam) |
| Distance formula | `distance = (real_hand_width * focal_length) / pixel_width`, with z-coordinate refinement |
| Optimal gesture recognition distance | 30-60cm (green), acceptable: 20-80cm (yellow) |
| Slide rendering | LibreOffice converts PPTX to PDF (150 DPI), then pdf2image converts to images with LANCZOS resampling |
| Speech recognition | Google Speech Recognition API with dynamic energy threshold and 0.8s pause threshold |
| Drawing layers | Separate NumPy arrays per slide for non-destructive annotations with 10 color options |
| Laser pointer | Red dot with 4-layer glow effect (sizes: 15, 12, 9, 6 pixels) |
| Eraser | 30-pixel radius with semi-transparent orange cursor and crosshair |
| Transcript display | YouTube-style subtitle overlay with 50 chars/line, max 2 lines, semi-transparent black background |
| Gesture cooldown | 0.5 seconds to prevent double-triggering (1 second for color change, slide navigation, and clear actions) |

## Notes

- Ensure good lighting for hand tracking
- Position hand clearly in camera view at optimal distance (30-60cm shown in green)
- Monitor the distance indicator in top-right corner of camera window for best gesture recognition
- Speech recognition requires internet connection
- LibreOffice and poppler are required for PPTX rendering
- Transcript appears as YouTube-style subtitles at bottom of presentation window
- All 10 colors available: Red, Green, Blue, Yellow, Cyan, Magenta, White, Orange, Purple, Black
- Gesture patterns shown in real-time on camera feed for easy reference
