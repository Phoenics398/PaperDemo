import os
from setuptools import setup, find_packages

# 使用绝对路径读取 requirements.txt
def get_requirements():
    req_path = os.path.join(os.path.dirname(__file__), 'requirements.txt')
    with open(req_path, 'r', encoding='utf-8') as f:
        return [
            line.strip() for line in f \
                if line.strip() and not line.startswith(('#', '-'))
        ]

setup(
    name='PaperDemo',
    version='0.1.0',
    packages=find_packages(exclude=[
        '.temp', '.backup', '.temps'
    ]),
    include_package_data=True,
    install_requires=get_requirements(),
    author='Leon HO',
    author_email='',  # 可以添加邮箱地址
    description='基于特定论文、著作、报告等模板生成内容的工具',
    long_description=open('readme.md', encoding='utf-8').read() if os.path.exists('readme.md') else '',
    long_description_content_type='text/markdown',
    url='',  # 可以添加项目URL
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',  # 选择合适的许可证
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.8',  # 指定Python版本要求
)