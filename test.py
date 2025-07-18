import psutil
process = psutil.Process()
mem_info = process.memory_info()
print(f"当前内存使用: {mem_info.rss / (1024 ** 2):.2f}MB")
