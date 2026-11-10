# WordEval website

The WordEval project page and its bilingual research article. This is a static
website; no evaluation backend is required.

## Local preview

From this directory:

```sh
python -m pip install -r tools/requirements.txt
python tools/preview.py
```

Open `http://127.0.0.1:8775/` for the project page, `/blog/` for the English
article, or `/blog/zh.html` for the Chinese article.
The preview rebuilds the Chinese page on startup and disables caching. After
editing article content, run `python tools/build_blog.py` and refresh the page.

## Editing

- `index.html`: project content, navigation, video, slides and poster links.
- `blog/index.html`: English article and shared figures.
- `blog/zh-CN.json`: Chinese text keyed to the English article.
- `blog/zh.html`: generated Chinese page; rebuild rather than editing directly.
- `web/`: project charts, task explorer, document renders and fonts.
- `assets/` and `blog/assets/`: styles and article interactions.
- `media/` and `slides.html`: presentation materials.

Images and chart labels are shared across the two article languages. All local
links are relative, so the site can be served at a domain root or under a project
path such as `/WordEval/`. Only this directory needs to be served as the website.

Repository licenses remain in effect. Font license notices are in `web/fonts/`.
