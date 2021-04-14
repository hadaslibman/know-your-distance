#!/bin/bash

brew list mvn || `brew update && brew install mvn`
mvn clean package
echo =============================================
echo =============================================
echo =============================================
echo =============================================
java -jar target/Distance-0.0.1-SNAPSHOT-jar-with-dependencies.jar "$1"
