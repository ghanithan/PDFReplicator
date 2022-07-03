#!/bin/sh
Echo "Starting Execution"
java -version
java -Xms4g -Xmx5g -jar ./build/libs/app-0.9-all.jar -verbose
pause