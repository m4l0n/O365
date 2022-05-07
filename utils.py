def clear_screen():
    """
    Clears the terminal/console screen.
    """
    print("\033[H\033[J", end = "")


def clear_last_input():
    print("\033[1A[\033[2K" + '\033[1A')