import os
from read_ppt import generate_html_from_pptx


def find_pptx_file(input_dir="input"):
    for file in os.listdir(input_dir):
        if file.endswith(".pptx"):
            return os.path.join(input_dir, file)
    return None


def main():
    # === 設定ここから ===
    pptx_path = find_pptx_file("input")
    template_path = "template/template.html"
    output_filename = "dayservice-aigi"  # 拡張子なし
    page_title = "デイサービスあいぎ"  # ページタイトルを設定
    start_slide = 24
    end_slide = 26
    # === 設定ここまで ===

    if not pptx_path:
        print("inputフォルダにpptxファイルがありません。")
        return

    # HTML断片を取得（パーツ名は自動判定に任せる）
    html_contents = generate_html_from_pptx(
        pptx_path,
        start_slide,
        end_slide,
        page_title,
        output_filename=output_filename,
    )

    # テンプレート読み込み
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    # 置換
    html = template.replace("{contents}", html_contents).replace(
        "{pagettl}", page_title
    )

    # 出力
    output_html = f"output/{output_filename}.html"
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"完了：{output_html} に出力しました！")


if __name__ == "__main__":
    main()
