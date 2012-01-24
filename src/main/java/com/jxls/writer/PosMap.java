package com.jxls.writer;

import java.util.ArrayList;
import java.util.List;

/**
 * Date: Nov 3, 2009
 *
 * @author Leonid Vysochyn
 */
public class PosMap {
    private List<Cell> destination = new ArrayList<Cell>();
    private Cell origin;

    public PosMap(Cell origin, List<Cell> destination) {
        this.origin = origin;
        this.destination = destination;
    }

    public PosMap(Cell origin) {
        this.origin = origin;
    }

    public PosMap() {
    }

    public void addToDestination(Cell cell){
        destination.add(cell);
    }

    public void setDestination(List<Cell> destination) {
        this.destination = destination;
    }

    public List<Cell> getDestination() {
        return destination;
    }

    public void combine(PosMap posMap) {
        if( posMap.isOneToOneMap() ){
            Size delta = posMap.getDelta();
            List<Cell> newDestination = new ArrayList<Cell>();
            if( destination.isEmpty() ){
                newDestination.add( posMap.getDestination().get(0) );
            }else{
                for (Cell cell : destination) {
                    newDestination.add(cell.add(delta));
                }
            }
            destination = newDestination;
        }
    }

    private Size getDelta() {
        return destination.get(0).minus(origin);
    }

    private boolean isOneToOneMap() {
        return destination.size()==1;
    }

    @Override
    public String toString() {
        return "Origin: " + origin + ", Destination: " + destination.toString();
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        PosMap posMap = (PosMap) o;

        if (destination != null ? !destination.equals(posMap.destination) : posMap.destination != null)
            return false;
        if (origin != null ? !origin.equals(posMap.origin) : posMap.origin != null) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = destination != null ? destination.hashCode() : 0;
        result = 31 * result + (origin != null ? origin.hashCode() : 0);
        return result;
    }
}
