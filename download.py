from roboflow import Roboflow
rf = Roboflow(api_key="KM4EAxzplNv5QNYs5JZC")
project = rf.workspace("helmetdetector-1ia1h").project("helmet-detection-using-yolo-v8-r4rgl")
version = project.version(1)
dataset = version.download("yolov8")