Welcome to this nested tables test
==================================

Table 1
-------

A normal table.

.. Hack to get tablestyle: DocxTableStyle Table Grid

+-------+-------+
| Hello | There |
+-------+-------+
| World | Q     |
+-------+-------+

Table 2
-------

A table with a list

.. Hack to get tablestyle: DocxTableStyle Table Grid

+-------+-------+
| Hello | There |
+-------+-------+
| World | - w1  |
|       | - w2  |
+-------+-------+

Table 3
-------

A table with a table

.. Hack to get tablestyle: DocxTableStyle Table Grid

+-------+---------+
| Hello | There   |
+-------+---------+
| World | +--+--+ |
|       | |A |B | |
|       | +--+--+ |
+-------+---------+

Table 4
-------

A multicol table with a multicol table

.. Hack to get tablestyle: DocxTableStyle Table Grid

+-------+---------+---+
| Hello | There   | H |
+-------+---------+---+
| World | +--+--+     |
|       | |A |B |     |
|       | +--+--+     |
|       | |AB   |     |
|       | +--+--+     |
+-------+-------------+

List 1
------

A normal list

 * Hello
 * World


List 2
------

A table in a list

 * Hello
 * +-------+-------+
   | Hello | There |
   +-------+-------+
   | World | Q     |
   +-------+-------+
 * There
