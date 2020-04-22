package com.example.demo.util;

import java.util.function.Supplier;
import java.util.stream.IntStream;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.stereotype.Component;

@Component
public class UploadUtil {
	/*public Supplier<Stream<Row>> getRowStreamSupplier(Iterable<Row> rows){
		return ()->getRowStream(rows);
	}*/
	public Supplier<Stream<Row>> getRowStreamSupplier(Iterable<Row> rows){
		return ()->getStream(rows);
	}
	
	public <T> Stream<T> getStream(Iterable<T> iterable){
		return StreamSupport.stream(iterable.spliterator(), false);
	}
	//we commented because we make above stream generic so two stream not needed
	/*public Stream<Row> getRowStream(Iterable<Row> rows){
		return StreamSupport.stream(rows.spliterator(), false);
	}*/
	
	//we commented because we make above strem generic so two stream not needed
	/*public Stream<Cell> getCellStream(Iterable<Cell> cells){
		return StreamSupport.stream(cells.spliterator(), false);*/
	public Supplier<Stream<Integer>> cellIteratorSupplier(int end) {
		return () -> numberStream(end);
	}

	public Stream<Integer> numberStream(int end) {
		return IntStream.range(0, end).boxed();
	}
}
