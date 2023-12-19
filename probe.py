import re

pattern = r"\b\w+\b\s\b[\w-]+\n"
print(re.search(pattern, "Vicki Herbert-Smith\n"))