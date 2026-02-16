# excelbench-rust (shim)

This is a compatibility shim for legacy code that imports `excelbench_rust`.

It re-exports the WolfXL native extension module (`wolfxl._rust`).

Install:

```bash
pip install excelbench-rust
```

Usage:

```python
import excelbench_rust

info = excelbench_rust.build_info()
print(info)
```
