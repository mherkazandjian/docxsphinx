# md_images — image embedding and sizing

## Plain image (native size)

![sample image](image1.png)

## Image with width and centered alignment (MyST attrs_block)

```{image} image1.png
:alt: A centered 2-inch-wide image
:width: 2in
:align: center
```

## Image with height in cm

```{image} image1.png
:alt: 3cm tall
:height: 3cm
```

## Image with width in pixels

```{image} image1.png
:alt: 200px wide
:width: 200px
```

## Inline image inside a paragraph

Regular text, then an inline image — ![inline](image1.png) — and more
text after it flowing as a single paragraph.

## Repeated image, native size

![first use](image1.png)
![second use](image1.png)
