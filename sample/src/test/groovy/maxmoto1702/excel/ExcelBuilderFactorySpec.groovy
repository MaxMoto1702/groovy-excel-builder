package maxmoto1702.excel

import spock.lang.*

class ExcelBuilderFactorySpec extends Specification {
    def "test get builder"() {
        setup:
        def factory = new ExcelBuilderFactory()

        when:
        def builder = factory.builder

        then:
        builder != null
    }
}
