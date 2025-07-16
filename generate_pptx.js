
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio'); // HTML解析ライブラリ

async function generatePresentation() {
    let pptx = new pptxgen();
    pptx.layout = 'LAYOUT_WIDE'; // ワイドスクリーン (16:9)

    const htmlFilePath = path.join(__dirname, 'index.html');
    const htmlContent = fs.readFileSync(htmlFilePath, 'utf8');
    const $ = cheerio.load(htmlContent);

    // 各セクションをスライドとして追加
    $('.reveal .slides section').each((index, element) => {
        const $section = $(element);
        let slide = pptx.addSlide();

        // スライド内のh1, h2, p, ul, li タグを抽出
        $section.find('h1, h2, h3, p, ul, li').each((i, el) => {
            const $el = $(el);
            const text = $el.text().trim();
            if (text) {
                let options = { x: 0.5, y: 0.5 + i * 0.5, w: 9, h: 0.5 };
                let textOptions = { fontSize: 18, color: '000000' }; // デフォルトは黒

                if (el.tagName === 'h1') {
                    textOptions.fontSize = 36;
                    options.y = 0.5;
                } else if (el.tagName === 'h2') {
                    textOptions.fontSize = 28;
                    options.y = 1.0 + i * 0.4;
                } else if (el.tagName === 'h3') {
                    textOptions.fontSize = 22;
                    options.y = 1.5 + i * 0.3;
                } else if (el.tagName === 'li') {
                    textOptions.bullet = true;
                    options.x = 0.7;
                    options.y = 0.5 + i * 0.3;
                }

                slide.addText(text, { ...options, ...textOptions });
            }
        });
    });

    // プレゼンテーションを保存
    const outputFilePath = path.join(__dirname, 'presentation.pptx');
    pptx.writeFile({ fileName: outputFilePath });
    console.log(`プレゼンテーションが ${outputFilePath} に生成されました。`);
}

generatePresentation();
