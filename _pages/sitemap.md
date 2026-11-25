---
layout: archive
title: "Sitemap"
permalink: /sitemap/
author_profile: true
---

<ul>
  <li><a href="{{ '/' | relative_url }}">About</a></li>
  {% for item in site.data.navigation.main %}
  <li><a href="{{ item.url | relative_url }}">{{ item.title | escape }}</a></li>
  {% endfor %}
</ul>

<h2>Research portfolio</h2>
{% assign projects = site.portfolio | sort: "order" %}
<ul>
  {% for project in projects %}
  <li><a href="{{ project.url | relative_url }}">{{ project.title | escape }}</a></li>
  {% endfor %}
</ul>

[XML sitemap]({{ '/sitemap.xml' | relative_url }})
