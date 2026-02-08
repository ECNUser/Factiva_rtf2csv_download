# Factiva_rtf2csv_download

该脚本用于转化从Factiva数据库下载的rtf文件转化为csv文件：

**使用示例**：

1. 处理单个文件：
   ```
   python factivartf2xlsx.py -i input.rtf
   ```

2. 处理目录中所有RTF文件：
   ```
   python factivartf2xlsx.py -i rtf_files/
   ```

3. 处理目录并合并到单个文件：
   ```
   python factivartf2xlsx.py -i rtf_files/ -o merged.csv -m
   ```

4. 处理目录并保存到指定输出目录：
   ```
   python factivartf2xlsx.py -i rtf_files/ -o output_csv/
   ```
        
