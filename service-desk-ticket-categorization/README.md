# Service Desk Ticket Categorization

The files in this folder relate to a project where the NPU on a Qualcomm Snapdragon X Elite processor is used to categorize incoming service desk tickets

## Project Status

This project is not complete. After working on the solution, we encountered an issue where we were able to output the ONNX model, but could not NPU-optimize it because the model was built using text data. The ML model needs to be refactored to use a "roll your own featurizer" that pre-converts text input to an integer array.
