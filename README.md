# datamine_python_script
using python to interact with datamine studioRM
The initial connection is established using;

import win32com.client
oDmApp = win32com.client.Dispatch("Datamine.StudioRM.Application")

This gives you access to datamine application to run commands and any advanced com operations.
