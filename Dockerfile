# 指定要使用的基本镜像
FROM tiangolo/uwsgi-nginx-flask:python3.6
# 将所有要用到的文件，文件夹复制到新的地址，注意原始地址与目标地址中间有一个空格
COPY requirements.txt /tmp/
COPY ./scripts /app
COPY ./data /data
# 设置该地址为工作地址
WORKDIR /app
# 更新pip
RUN pip install -i https://pypi.tuna.tsinghua.edu.cn/simple/ -U pip
# 安装所有要用到的外部包的镜像
RUN pip install -i https://pypi.tuna.tsinghua.edu.cn/simple/ -r /tmp/requirements.txt
# 设置entry point，由于python版本为3.6， entry point设置为python3
ENTRYPOINT ["python3"]
# 运行主程序
CMD ["main.py"]