"""VBA arrays with inclusive bounds, assignment copies and Preserve."""

from itertools import product

from .errors import VBAError
from .values import MISSING, coerce, default_for, integer


class VBAArray:
    _vba_value_type = True

    def __init__(self, bounds=None, type_name="Variant", *, fixed=False, fixed_length=None):
        self.bounds = None if bounds is None else tuple((integer(a), integer(b)) for a, b in bounds)
        self.type_name = type_name
        self.fixed = fixed
        self.fixed_length = fixed_length
        self._data = {}
        if self.bounds is not None and (not self.bounds or len(self.bounds) > 60):
            raise VBAError(9)

    @classmethod
    def from_values(cls, values, base=0, type_name="Variant"):
        values = tuple(values)
        result = cls(((base, base + len(values) - 1),), type_name)
        for idx, item in enumerate(values, base):
            result.set((idx,), item)
        return result

    @property
    def _vba_type_name(self):
        return f"{self.type_name}()"

    def _index(self, indices):
        if self.bounds is None or len(indices) != len(self.bounds):
            raise VBAError(9)
        index = tuple(integer(value) for value in indices)
        if any(value < lower or value > upper for value, (lower, upper) in zip(index, self.bounds)):
            raise VBAError(9)
        return index

    def get(self, indices):
        index = self._index(indices)
        if index not in self._data:
            return default_for(self.type_name, self.fixed_length)
        return self._data[index]

    def set(self, indices, value):
        self._data[self._index(indices)] = coerce(value, self.type_name, self.fixed_length)

    def __iter__(self):
        if self.bounds is None:
            raise VBAError(9)
        for reversed_indices in product(*(range(lo, hi + 1) for lo, hi in reversed(self.bounds))):
            yield self.get(tuple(reversed(reversed_indices)))

    def copy(self):
        result = VBAArray(self.bounds, self.type_name, fixed=self.fixed, fixed_length=self.fixed_length)
        result._data = {key: value.copy() if isinstance(value, VBAArray) else value for key, value in self._data.items()}
        return result

    def redim(self, bounds, preserve=False, type_name=None):
        if self.fixed:
            raise VBAError(10)
        if any(upper < lower for lower, upper in bounds):
            raise VBAError(9)
        result = VBAArray(bounds, type_name or self.type_name, fixed_length=self.fixed_length)
        if preserve and self.bounds is not None:
            if (self.bounds[:-1] != result.bounds[:-1] or self.bounds[-1][0] != result.bounds[-1][0]
                    or self.type_name.casefold() != result.type_name.casefold()):
                raise VBAError(9)
            result._data = {key: value for key, value in self._data.items() if key[-1] <= result.bounds[-1][1]}
        return result

    def erase(self):
        self._data.clear()
        if not self.fixed:
            self.bounds = None

    def bound(self, dimension=MISSING, upper=False):
        index = 1 if dimension is MISSING else integer(dimension)
        if self.bounds is None or not 1 <= index <= len(self.bounds):
            raise VBAError(9)
        return self.bounds[index - 1][int(upper)]
