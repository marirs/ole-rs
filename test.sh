#!/bin/bash

# Build the workspace first.
cargo build

echo "RUNNING THE OLEID TOOL"
for file in data/*; do 
    echo "$file"; 
    "target/debug/oleid" --file "$file"
done

echo "RUNNING THE OLEOBJ TOOL"
for file in data/*; do 
    echo "$file"; 
    "target/debug/oleobj" --file "$file"
done
