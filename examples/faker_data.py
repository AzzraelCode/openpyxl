from random import randrange

import faker


def data_samples():
    fk = faker.Factory.create('ru_RU')
    return [[fk.name(), fk.phone_number(), randrange(0, 9)] for _ in range(100)]