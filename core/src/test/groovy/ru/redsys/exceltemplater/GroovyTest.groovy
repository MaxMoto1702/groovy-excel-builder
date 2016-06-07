package ru.redsys.exceltemplater

import org.junit.Test

class GroovyTest {
    def method() {
        method(null, null)
    }

    def method(value) {
        method([value: value], null)
    }

    def method(Map params) {
        method(params, null)
    }

    def method(Closure closure) {
        method(null, closure)
    }

    def method(Map params, Closure closure) {
        if (params?.value && closure) return closure.call(params.value)
        if (!params?.value && closure) return closure()
        if (params?.value && !closure) return params.value
    }

    @Test
    void test() {
        assert "test" == method(value: "test") { it }
        assert "test" == method() { "test" }
        assert "test" == method { "test" }
        assert "test" == method(value: "test")
        assert "test" == method("test")
        assert null == method()
    }

    @Test
    void testCollisions() {
        def a = ['a', 'b', 'c']
        def b = ['b', 'c', 'f']
        a.metaClass.methods.each {
            println "$it"
        }
        assert ['b', 'c'] == b.intersect( a )
    }
}
