#!/bin/sh
Echo "Starting Execution"
java -version
java -Xms4g -Xmx5g -jar ./build/libs/app-1.0-all.jar -verbose
