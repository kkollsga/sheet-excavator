# 1 - Set up parallell processing with tokio

# 2 - Set up basic excel get cell function with calamine

# 3 - Set up python test

# 4 - Build out complex excel data extraction
```python
extraction_details = [
    {"method":"single", "sheet":"Sheet 1", "cell":"a2", "value_name": "dataname"},
    {"method":"single", "sheet":"Sheet 1", "column":1, "row":6, "value_name":"description"},
    {"method":"multirow","startrow":5,"endrow":10,"filter":{"col":1,"not empty"}},
    ...
]
```
