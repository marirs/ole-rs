#!/bin/bash

for file in data/*; do 
    echo "$file"; 
    cargo run --example ole_cli --features="blocking" -- "$file" >"$file".res
done