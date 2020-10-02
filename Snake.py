# 'U' : - For Up
# 'J' : - For Down
# 'H' : - For Left
# 'K' : - For Right


import  turtle
import time
import random

delay = 0.1

#Scores

score = 0       # initial score
h_score = 0     # initial high score

# Scree Setup

win = turtle.Screen()
win.title("Snake Game by Code with Aj")
win.bgcolor("black")
win.setup(width=600, height=600)
win.tracer(0)   # Turn off the screen updates

# Define head of snake
head = turtle.Turtle()
head.speed(0)
head.shape("triangle")
head.color("white")
head.penup()
head.goto(0,0)
head.direction = "stop"

#Food

food =turtle.Turtle()
food.speed(0)
food.shape("circle")
food.color("red")
food.penup()
food.goto(0,100)

segment = []

#Pen

pen = turtle.Turtle()
pen.speed(0)
pen.shape("square")
pen.color("yellow")
pen.penup()
pen.hideturtle()
pen.goto(0,260)
pen.write("Score: 0 High Score: 0 ", align="center",font=("Courier",24,"normal"))


#---------------------Functions---------------------------------

def up():
    if head.direction != "down":
        head.direction ="up"

def down():
    if head.direction != "up":
        head.direction ="down"

def left():
    if head.direction != "right":
        head.direction ="left"

def right():
    if head.direction != "left":
        head.direction ="right"

def movement():
    if head.direction == "up":
        y=head.ycor()
        head.sety(y+20)

    if head.direction == "down":
        y=head.ycor()
        head.sety(y-20)
        
    if head.direction == "right":
        x=head.xcor()
        head.setx(x+20)
    
    if head.direction == "left":
        x=head.xcor()
        head.setx(x-20)

# Keyboard 

win.listen()
win.onkeypress(up,"u")
win.onkeypress(down,"j")
win.onkeypress(left,"h")
win.onkeypress(right,"k")

# Main Loop

while True:
    win.update()

    # check Collision
    if head.xcor()>290 or head.xcor()<-290 or head.ycor()>290 or head.ycor()<-290:
        time.sleep(1)
        head.goto(0,0)
        head.direction = "stop"

        # hide segment
        for seg in segment:
            seg.goto(1000,1000)

        # Clear Segment list
        segment.clear()

        #Reset Score
        score = 0

        # Reset  delay
        delay = 0.1

        pen.clear()
        pen.write("Score:{} High Score: {}".format(score,h_score), align="center",font=("times new roman ",24,"normal"))

        #check collision with food

    if head.distance(food)<20:
        #Move food to random place
        x = random.randint(-290,290)
        y = random.randint(-290,290)
        food.goto(x,y)

        #Add segment

        new = turtle.Turtle()
        new.speed(0)
        new.shape("square")
        new.color("light blue")
        new.penup()
        segment.append(new)

        # Short the delay
        delay -= 0.001

        # Increase the score
        score +=10

        if score > h_score:
            h_score = score

        pen.clear()
        pen.write("Score: {} High Score: {}".format(score,h_score),align="center", font=("times new roman",24,"normal"))

    # Move the end segment first in reverse order
    for index in range(len(segment)-1,0,-1):
        x = segment[index-1].xcor()
        y = segment[index-1].ycor()
        segment[index].goto(x,y)

    #Move segment to o
    if len(segment) > 0:
        x = head.xcor()
        y = head.ycor()
        segment[0].goto(x,y)

    movement()

    # Check head collision with body

    for seg in segment:
        if seg.distance(head) < 20:
            time.sleep(1)
            head.goto(0,0)
            head.direction = "stop"

            # hide segment
            for seg in segment:
                seg.goto(1000,1000)

            #Clear
            segment.clear()

            # Reset Score 

            score = 0 
            delay = 0.1

            pen.clear()
            pen.write("Score:{} High Score: {}".format(score,h_score),align="center",font=("times new roman",24,"normal"))

    time.sleep(delay)
win.mainloop()

        
                
