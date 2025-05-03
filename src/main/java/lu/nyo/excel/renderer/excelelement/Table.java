package lu.nyo.excel.renderer.excelelement;

import lu.nyo.excel.renderer.Renderable;

import java.util.stream.Stream;

public interface Table extends Renderable {

    void tableContentInitializer(Holder headerHolder,
                                 Holder bodyHolder,
                                 Holder footerHolder);

    @Override
    default Object getRenderingEngineSelector() {
        return Table.class;
    }

    class Holder {
        Stream<Row> rowStream;

        public Holder(Stream<Row> rowStream) {
            this.rowStream = rowStream;
        }

        public Stream<Row> get() {
            return rowStream;
        }

        public void set(Stream<Row> rowStream) {
            this.rowStream = rowStream;
        }
    }
}
