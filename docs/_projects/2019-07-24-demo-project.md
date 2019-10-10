---
title: Mobile app
subtitle: High-fidelity mobile app designs for a super awesome social media company.
date: 2019-07-24 00:00:00
description: This is a demo post that shows what you can do inside portfolio and blog posts. We’ve included everything you need to create engaging posts and case studies to show off your work in a beautiful way.
featured_image: demo.jpg
accent_color: '#4C60E6'
gallery_images:
  - demo.jpg
  - demo.jpg
  - demo.jpg
---

This page is a demo that shows everything you can do inside portfolio and blog posts.

We've included everything you need to create engaging posts about your work, and show off your case studies in a beautiful way.

![]({{ '/images/demo.jpg' | relative_url }})

**Obviously,** we’ve styled up *all the basic* text formatting options [available in markdown](https://github.com/adam-p/markdown-here/wiki/Markdown-Cheatsheet).

You can create lists:

* Simple bulleted lists
* Like this one
* Are cool

And:

1. Numbered lists
2. Like this other one
3. Are great too

You can also add blockquotes, which are shown at a larger width to help break up the layout and draw attention to key parts of your content:

> “Simple can be harder than complex: You have to work hard to get your thinking clean to make it simple. But it’s worth it in the end because once you get there, you can move mountains.”

The theme also supports markdown tables:

| Item                 | Author        | Supports tables? | Price |
|----------------------|---------------|------------------|-------|
| Duet Jekyll Theme    | Jekyll Themes | Yes              | $49   |
| Index Jekyll Theme   | Jekyll Themes | Yes              | $49   |
| Journal Jekyll Theme | Jekyll Themes | Yes              | $49   |

You can throw in some horizontal rules too:

---

#### Image galleries

Here's a really neat custom feature we added – galleries:

{% include post-components/gallery.html
	columns = 2
	full_width = true
	images = "/images/demo.jpg,/images/demo.jpg,/images/demo.jpg"
%}

Inspired by the Galleries feature from WordPress, we've made it easy to create grid layouts for your images. Just use a simple Liquid snippet in your post to create a masonry grid image layout:

{% raw %}
```liquid
{% include post-components/gallery.html
	columns = 2
	full_width = true
	images = "/images/demo.jpg,/images/demo.jpg,/images/demo.jpg,/images/demo.jpg"
%}
```
{% endraw %}

*See what we did there? Code and syntax highlighting is built-in too!*

Change the number inside the 'columns' setting to create different types of gallery for all kinds of purposes. You can even click on each image to seamlessly enlarge it on the page.


#### Image carousels

Here's another gallery with only one column, which creates a carousel slide-show instead.

A nice little feature: the carousel only advances when it is in view, so your visitors won't scroll down to find it half way through your images.

{% include post-components/gallery.html
	columns = 1
	full_width = true
	images = "/images/demo.jpg,/images/demo.jpg,/images/demo.jpg"
%}

#### What about videos?

Videos are an awesome way to show off your work in a more engaging and personal way, and we’ve made sure they work great on our themes. Just paste an embed code from YouTube or Vimeo, and the theme makes sure it displays perfectly:

{% include post-components/video.html
	url = "https://player.vimeo.com/video/270725085?color=6c6e95&title=0&byline=0"
	full_width = true
%}

### Pretty cool, huh?

We've packed this theme with powerful features to show off your work.
Why not put them to use on your new website?

<a href="https://jekyllthemes.io/theme/made-portfolio-jekyll-theme" class="button--fill">Get This Theme</a>