<!--#include file="../class_testFixture/testFixture.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	tf.assertEqual lib.range(1, 3, 1), array(1, 2, 3), "lib.range"
	tf.assertEqual lib.range(5, 0, -1), array(5, 4, 3, 2, 1, 0), "lib.range"
	tf.assertEqual lib.range(1, 3, 0.5), array(1, 1.5, 2, 2.5, 3), "lib.range"
	tf.assertEqual lib.range(1, 3.2, 0.5), array(1, 1.5, 2, 2.5, 3), "lib.range"
	tf.assertEqual lib.range(1, 1.2, 1), array(1), "lib.range"
	tf.assertEqual lib.range(1, 2, 1), array(1, 2), "lib.range"
	tf.assertHas lib.range(1, 20, 1), 1, "lib.range"
	tf.assertHas lib.range(20, 1, -1), 10, "lib.range"
	tf.assertHas lib.range(1, 20, 1), 20, "lib.range"
	tf.assertHas lib.range(1, 20, 1), 15, "lib.range"
	tf.assertHas lib.range(-10, 0, 1), 0, "lib.range"
	tf.assertHas lib.range(-10, 0, 1), -10, "lib.range"
	tf.assertHas lib.range(-10, 0, 1), -5, "lib.range"
	tf.assertHas lib.range(-10, 10, 1), 10, "lib.range"
	tf.assertHas lib.range(0, 100, 0.5), 10, "lib.range"
	tf.assertHas lib.range(0, 100, 0.5), 50.5, "lib.range"
	tf.assertHas lib.range(0, 100, 0.5), 100, "lib.range"
	tf.assertHas lib.range(0, 100, 0.5), 0, "lib.range"
	tf.assertHasNot lib.range(0, 100, 0.5), 100.5, "lib.range"
	
	'TODO: michal: lib range is not working with floats perfectly. I am leaving it for now because with int it works fine.
	'i dunno now why it has a little offset after sometime.. uncomment next line to see the problem:
	'str.write(str.arrayToString(lib.range(-10, -5, 0.1), " - "))
	tf.assertHas lib.range(-10, -5, 0.1), -5.2, "lib.range"
	tf.assertHas lib.range(1, 20, 0.1), 9.6, "lib.range"
	tf.assertHas lib.range(1, 20, 0.1), 20, "lib.range"
	tf.assertHas lib.range(1, 20, 0.1), 17.6, "lib.range"
end sub
%>