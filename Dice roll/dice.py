class Dice:
    def roll(self):
        import random
        x = random.randint(1, 6)
        y = random.randint(1, 6)
        return x, y