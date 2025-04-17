# 使用阿里云镜像源的Python 3.10作为基础镜像
FROM python:3.10-slim

# 设置工作目录
WORKDIR /app

# 设置pip的清华镜像源并安装uv
RUN pip install -i https://pypi.tuna.tsinghua.edu.cn/simple pip --upgrade && \
    pip install -i https://pypi.tuna.tsinghua.edu.cn/simple uv

# 复制当前项目到容器
COPY . /app/

# 创建虚拟环境并安装依赖
RUN uv venv && \
    . .venv/bin/activate && \
    uv pip install -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple -e .

# 设置环境变量
ENV FASTMCP_PORT=8000
ENV EXCEL_FILES_PATH=/app/excel_files

# 创建excel文件目录
RUN mkdir -p /app/excel_files

# 暴露端口
EXPOSE 8000

# 启动服务
CMD ["uv", "run", "excel-mcp-server"] 
