package com.javacook.easyexcelaccess;

import com.javacook.coordinate.CoordinateInterface;
import com.javacook.coordinate.sequencer.*;

import static com.javacook.easyexcelaccess.ExcelCoordinate.COL_MAX;
import static com.javacook.easyexcelaccess.ExcelCoordinate.ROW_MAX;

/**
 * A coordinate iterator which visits a series of Excel row or columns cells for example:
 * <pre>
 *     ExcelCoordinateAccessor excel = new ExcelCoordinateAccessor(file, sheet);
 *     new ExcelCoordinateSequencer()
 *         .forCol('C')
 *         .fromRow(3).toRow(9)
 *         .forEach(coord -> System.out.println(excel.read(coord)));
 * </pre>
 */
public class ExcelCoordinateSequencer {

    // Deligate
    final protected CoordinateSequencer<ExcelCoordinate> cs;

    public ExcelCoordinateSequencer() {
        cs = new CoordinateSequencer<>(
                (x, y) -> new ExcelCoordinate(x+1, y+1)
        );
    }

    private ExcelCoordinateSequencer fromX(int x) {
        cs.fromX(x);
        return this;
    }

    private ExcelCoordinateSequencer fromY(int y) {
        cs.fromY(y);
        return this;
    }

    public ExcelCoordinateSequencer from(int col, int row) {
        return fromCol(col).fromRow(row);
    }

    public ExcelCoordinateSequencer from(ExcelCoordinate coord) {
        return from(coord.col(), coord.row());
    }

    public ExcelCoordinateSequencer from(CoordinateInterface coord) {
        cs.from(coord);
        return this;
    }

    private ExcelCoordinateSequencer toX(int xTo) {
        cs.toX(xTo);
        return this;
    }

    private ExcelCoordinateSequencer toY(int yTo) {
        cs.toY(yTo);
        return this;
    }

    public ExcelCoordinateSequencer to(ExcelCoordinate coord) {
        return to(coord.col(), coord.row());
    }

    public ExcelCoordinateSequencer to(CoordinateInterface c) {
        cs.to(c);
        return this;
    }

    public ExcelCoordinateSequencer to(int col, int row) {
        return toCol(col).toRow(row);
    }


    private ExcelCoordinateSequencer forX(int x) {
        cs.forX(x);
        return this;
    }

    private ExcelCoordinateSequencer forY(int y) {
        cs.forY(y);
        return this;
    }

    public ExcelCoordinateSequencer forCoord(CoordinateInterface c) {
        return forX(c.x()).forY(c.y());
    }

    private ExcelCoordinateSequencer lenX(int len) {
        cs.lenX(len);
        return this;
    }

    public ExcelCoordinateSequencer lenY(int len) {
        cs.lenY(len);
        return this;
    }

    public ExcelCoordinateSequencer width(int width) {
        return lenX(width);
    }

    public ExcelCoordinateSequencer height(int height) {
        return lenY(height);
    }


//    public ExcelCoordinateSequencer maxWidth() {
//        return width(COL_MAX);
//    }
//
//    public ExcelCoordinateSequencer maxHeight() {
//        return height(ROW_MAX);
//    }



    private ExcelCoordinateSequencer stepX(int step) {
        cs.stepX(step);
        return this;
    }

    private ExcelCoordinateSequencer stepY(int step) {
        cs.stepY(step);
        return this;
    }

    public ExcelCoordinateSequencer enter() {
        cs.enter();
        return this;
    }

    public ExcelCoordinateSequencer sequence() {
        cs.sequence();
        return this;
    }

    public ExcelCoordinateSequencer forEach(Consumer<ExcelCoordinate> action) {
        cs.forEach(action);
        return this;
    }

    public ExcelCoordinateSequencer forEach(ConsumerAndCounter<ExcelCoordinate> action) {
        cs.forEach(action);
        return this;
    }

    public ExcelCoordinateSequencer forEachPair(PairConsumer<ExcelCoordinate> action) {
        cs.forEachPair(action);
        return this;
    }

    public ExcelCoordinateSequencer forEachPair(PairConsumerAndCounter<ExcelCoordinate> action) {
        cs.forEachPair(action);
        return this;
    }


    public ExcelCoordinateSequencer forEachArray(ArrayConsumer<ExcelCoordinate> action) {
        cs.forEachArray(action);
        return this;
    }

    public ExcelCoordinateSequencer forEachArray(ArrayConsumerAndCounter<ExcelCoordinate> action) {
        cs.forEachArray(action);
        return this;
    }


    public ExcelCoordinateSequencer stopWhen(Predicate<ExcelCoordinate> action) {
        cs.stopWhen(action);
        return this;
    }

    public ExcelCoordinateSequencer stopWhen(PredicateAndCounter<ExcelCoordinate> action) {
        cs.stopWhen(action);
        return this;
    }

    public ExcelCoordinateSequencer stopWhenPair(PairPredicate<ExcelCoordinate> action) {
        cs.stopWhenPair(action);
        return this;
    }

    public ExcelCoordinateSequencer stopWhenPair(PairPredicateAndCounter<ExcelCoordinate> action) {
        cs.stopWhenPair(action);
        return this;
    }


    public ExcelCoordinateSequencer stopWhenArray(ArrayPredicate<ExcelCoordinate> action) {
        cs.stopWhenArray(action);
        return this;
    }

    public ExcelCoordinateSequencer stopWhenArray(ArrayPredicateAndCounter<ExcelCoordinate> action) {
        cs.stopWhenArray(action);
        return this;
    }

    public ExcelCoordinateSequencer fromRow(int row) {
        return fromY(row - 1);
    }

    public ExcelCoordinateSequencer toRow(int row) {
        return toY(row);
    }

    public ExcelCoordinateSequencer toRowMax() {
        return toRow(ROW_MAX);
    }

    public ExcelCoordinateSequencer forRow(int row) {
        return forY(row - 1);
    }

    public ExcelCoordinateSequencer fromCol(int col) {
        return fromX(col - 1);
    }

    public ExcelCoordinateSequencer fromCol(char col) {
        return fromCol(Character.toString(col));
    }

    public ExcelCoordinateSequencer fromCol(String col) {
        return fromCol(ExcelUtil.calculateColNo(col));
    }

    public ExcelCoordinateSequencer toCol(int col) {
        return toX(col);
    }

    public ExcelCoordinateSequencer toCol(char col) {
        return toCol(Character.toString(col));
    }

    public ExcelCoordinateSequencer toCol(String col) {
        return toCol(ExcelUtil.calculateColNo(col));
    }

    public ExcelCoordinateSequencer toColMax() {
        return toCol(COL_MAX);
    }

    public ExcelCoordinateSequencer forCol(int col) {
        return forX(col - 1);
    }

    public ExcelCoordinateSequencer forCol(char col) {
        return forCol(Character.toString(col));
    }

    public ExcelCoordinateSequencer forCol(String col) {
        return forCol(ExcelUtil.calculateColNo(col));
    }

    public ExcelCoordinateSequencer horStep(int stepX) {
        return stepX(stepX);
    }

    public ExcelCoordinateSequencer vertStep(int stepY) {
        return stepY(stepY);
    }


}