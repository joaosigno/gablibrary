<!--#include file="../class_testFixture/testFixture.asp"-->
<!--#include file="dates.asp"-->
<%
set tf = new TestFixture
tf.run()

sub test_1()
	set dat = new Dates
	tf.assertEqual dat.weekOfTheYear(dateSerial(2008, 1, 1)), 1, "dat.weekOfTheYear"
	tf.assertEqual dat.weekOfTheYear(dateSerial(2008, 4, 7)), 15, "dat.weekOfTheYear"
	tf.assertEqual dat.weekOfTheYear(dateSerial(2008, 12, 31)), 1, "dat.weekOfTheYear"
	tf.assertEqual dat.weekOfTheYear(dateSerial(2008, 12, 28)), 52, "dat.weekOfTheYear"
	tf.assertEqual dat.weekOfTheYear(dateSerial(2010, 1, 3)), 53, "dat.weekOfTheYear"
	tf.assertEqual dat.weekOfTheYear(dateSerial(2008, 12, 29)), 1, "dat.weekOfTheYear"
	tf.assertEqual dat.weekOfTheYear(dateSerial(2009, 7, 30)), 31, "dat.weekOfTheYear"
end sub
%>