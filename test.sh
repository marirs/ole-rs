#!/bin/bash

echo "RUNNING THE OLEID TOOL"
for file in data/*; do 
    echo "$file"; 
    cargo run -- --file "$file"
done