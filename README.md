# Balon_Sensor_Control


# Capabilities:
- Scan for BLE Balon device. 
- Verify if it has already been used. 
If it has been used perform a rescan. If there is an issue check using an NFC scanner and do a manual input of the address
- Connect to device + subscribe to get sensor values
- Start reading Data (indefinitely), which will be displayed on a Graph
- Stop recording Data
- Store Data if the Test went fine
- Do a Re-Read if the Test was not satisfying
- It is possible to store data multiple times for the same device 
-> will be stored in different sheets but in the same Excel file the naming will be as follows: MACADDRESS_SERIALNUMBER_XX 
(xx holds the latest version to indicate there are more sheets in the file)
- Disconnect from Device


# TODO

- do error handling!!!
- ensure rescan + reconnecting is working
- gui not displaying properly when being killed off
- slight delay between pressing and values
