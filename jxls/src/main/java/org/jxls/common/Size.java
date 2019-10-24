package org.jxls.common;

/**
 * @author Leonid Vysochyn
 */
public class Size {
    public static final Size ZERO_SIZE = new Size(0,0);
    private int width;
    private int height;

    public Size() {
    }

    public Size(int width, int height) {
        this.width = width;
        this.height = height;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public int getWidth() {
        return width;
    }

    public int getHeight() {
        return height;
    }

    @Override
    public String toString() {
        return "(" + width + "," + height + ")";
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        Size size = (Size) o;
        return height == size.height && width == size.width;
    }

    @Override
    public int hashCode() {
        int result = width;
        result = 31 * result + height;
        return result;
    }

    public Size minus(Size aSize) {
        return new Size(width - aSize.getWidth(), height - aSize.getHeight());
    }

    public Size add(Size aSize) {
        return new Size(width + aSize.getWidth(), height + aSize.getHeight());
    }

    public Size append(int width, int height) {
        this.width += width;
        this.height += height;
        return this;
    }

    public Size append(Size aSize) {
        return append(aSize.getWidth(), aSize.getHeight());
    }
}
