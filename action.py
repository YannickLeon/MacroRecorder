class action:

    keycode: int
    pressed: float
    released: float
    duration: float
    x: int
    y: int
    state: int # 0: ready 1:pressed 2:done

    def __init__ (self,  keycode, pressed, released, x, y):
        self.keycode = keycode
        self.pressed = pressed
        self.released = released
        self.duration = released - pressed
        self.x = x
        self.y = y
        self.completed = 0