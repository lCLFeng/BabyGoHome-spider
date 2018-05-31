import json


def write_to_file(content):
    with open('FailHref.txt', 'a+', encoding='utf-8') as f:
        f.write(json.dumps(content, ensure_ascii=False)+'\n')
        f.close()


if __name__ == '__main__':
    write_to_file()

