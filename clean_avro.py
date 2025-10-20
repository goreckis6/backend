#!/usr/bin/env python3
"""Script to remove AVRO routes from server.ts"""

import re

# Read the file
with open('src/server.ts', 'r', encoding='utf-8') as f:
    content = f.read()

# Split into lines for processing
lines = content.split('\n')

# Find and remove AVRO route blocks
i = 0
new_lines = []
skip_until_closing = 0

while i < len(lines):
    line = lines[i]
    
    # Check if this is an AVRO route definition
    if ('AVRO' in line and 'Route:' in line) or \
       ('/convert/avro-' in line or '/convert/csv-to-avro' in line) and 'app.post' in line:
        # Skip until we find the closing }); for this route
        skip_until_closing = 1
        i += 1
        continue
    
    # If we're skipping, look for the closing
    if skip_until_closing > 0:
        # Count braces to find the end of the route handler
        if '{' in line:
            skip_until_closing += line.count('{')
        if '}' in line:
            skip_until_closing -= line.count('}')
        
        # If we've closed all braces and see });, we're done
        if skip_until_closing <= 0 and '});' in line:
            skip_until_closing = 0
            i += 1
            continue
        
        i += 1
        continue
    
    # Keep this line
    new_lines.append(line)
    i += 1

# Join back
new_content = '\n'.join(new_lines)

# Write back
with open('src/server.ts', 'w', encoding='utf-8') as f:
    f.write(new_content)

print("✅ AVRO routes removed successfully!")
print(f"Original lines: {len(lines)}")
print(f"New lines: {len(new_lines)}")
print(f"Lines removed: {len(lines) - len(new_lines)}")


