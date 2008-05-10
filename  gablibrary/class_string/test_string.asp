<!--#include virtual="/gab_Library/class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.debug = true
tf.run()

sub test_1()
	tf.assert str.matching("hello", "h.ll.", true), "str.matching"
end sub

sub test_2()
	tf.assertEqual "ein 1 zwei {1}", str.format("ein {0} zwei {1}", 1), "str.format does not work"
	tf.assertEqual "ein 1 zwei 2", str.format("ein {0} zwei {1}", array(1, 2)), "str.format does not work"
	tf.assertEqual "ein {0} zwei {1}", str.format("ein {0} zwei {1}", array()), "str.format does not work"
end sub
%>
