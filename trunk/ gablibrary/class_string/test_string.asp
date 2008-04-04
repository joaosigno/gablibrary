<!--#include virtual="/gab_Library/class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	tf.assert str.matching("hello", "h.ll.", true), "str.matching"
end sub
%>
