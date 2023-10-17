class RandomId():

    @staticmethod
    def get(num_digits=9):
        from random import randrange
        return randrange(10**(num_digits-1), 10**(num_digits)-1)
