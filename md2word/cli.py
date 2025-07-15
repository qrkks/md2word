import argparse

def main():
    parser = argparse.ArgumentParser(description='Markdown 转 Word 工具')
    parser.add_argument('input', help='输入的 Markdown 文件')
    parser.add_argument('-o', '--output', help='输出的 Word 文件', required=True)
    args = parser.parse_args()
    # TODO: 调用转换和宏处理流程
    print(f"输入: {args.input}，输出: {args.output}")

if __name__ == '__main__':
    main() 