"""
This is my module.
"""


class MyClass(object):
    """
    This is my class.
    """

    def with_function(self):
        """
        With a function

        :return: MyClass
        """
        return self


def just_a_function(x, y=4):
    """
    Just a function

    :param x: First number
    :param y: Second number
    :type x: int
    :type y: int
    :return: multiplication
    :rtype: int
    """
    return x * y
