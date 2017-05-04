import six

EMPTY_CELL = ''

class ConfigurationError(Exception):
    pass


class DoesntMatchException(Exception):
    pass


def str_or_none(o):
    return o is None or isinstance(o, six.string_types)


def numrow(s):
    result = 0
    for i in s.strip().upper():
        result = result * 26 + ord(i) - 64
    return result


def instantiate_if_class(cls_or_inst, cls, **kwargs):
    if isinstance(cls_or_inst, cls):
        return cls_or_inst
    result = cls_or_inst(**kwargs)
    if not isinstance(result, cls):
        raise ConfigurationError("Expected %s, got %s" % (cls.__name__, type(result)))
    return result


def instantiate_if_class_lst(lst, cls, **kwargs):
    return [instantiate_if_class(c, cls, **kwargs) for c in lst]

