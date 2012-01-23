package com.jxls.writer;

/**
 * Date: Mar 13, 2009
 *
 * @author Leonid Vysochyn
 */
public class Pos {
    int y;
    int x;

    public Pos(int x, int y) {
        this.x = x;
        this.y = y;
    }

    public int getY() {
        return y;
    }

    public int getX() {
        return x;
    }

    public Pos add(Size size){
        return new Pos(x + size.getWidth(), y + size.getHeight());
    }

    public Pos append(Size size){
        x += size.getWidth();
        y += size.getHeight();
        return this;
    }

    @Override
    public String toString() {
        return "(" + x + "," + y + ")";
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Pos pos = (Pos) o;

        if (y != pos.y) return false;
        if (x != pos.x) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = y;
        result = 31 * result + x;
        return result;
    }

    public boolean isLower(Pos pos) {
        return y > pos.getY();
    }

    public boolean isOnTheRightOf(Pos pos){
        return x > pos.getX();
    }

    public Size minus(Pos pos) {
        return new Size(x - pos.getX(), y - pos.getY());
    }

    public Pos addYModification(int yChange) {
        return new Pos(x, y + yChange);
    }

    public Pos addXModification(int xChange) {
        return new Pos(x + xChange, y);
    }
}
