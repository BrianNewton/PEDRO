import matplotlib.pyplot as plt
import matplotlib.lines as lines

class draggable_lines:
    def __init__(self, ax, start_coordinate, x_bounds, y_bounds):
        self.ax = ax
        self.c = ax.get_figure().canvas
        self.x_bounds = x_bounds
        self.y_bounds = y_bounds
        self.press = None

        self.line = lines.Line2D([start_coordinate, start_coordinate], y_bounds, color='r', picker=5)

        self.ax.add_line(self.line)
        self.c.draw_idle()
        self.sid = self.c.mpl_connect('button_press_event', self.on_press)
        self.sid = self.c.mpl_connect('motion_notify_event', self.on_motion)
        self.sid = self.c.mpl_connect('button_release_event', self.on_release)


    def on_press(self, event):
        if abs(event.xdata - self.line.get_xdata()[0]) < 3:
            self.press = (self.line.get_xdata()[0], event.xdata)
        return
    
    def on_motion(self, event):
        if self.press == None:
            return 
        try: 
            x0, xpress = self.press
            dx = event.xdata - xpress
            if self.x_bounds[0] >= (x0 + dx):
                self.line.set_xdata([self.x_bounds[0],self.x_bounds[0]])
            elif (x0 + dx) >= self.x_bounds[1]:
                self.line.set_xdata([self.x_bounds[1], self.x_bounds[1]])
            else:
                self.line.set_xdata([x0+dx, x0+dx])
        except:
            pass
    
    def on_release(self, event):
        self.press = None
        self.c.draw_idle()




def linear_regression(X, Y):
    # simple linear regression
    sumX = sum(X)
    sumY = sum(Y)
    meanX = sumX/len(X)
    meanY = sumY/len(Y)

    SSx = 0 # sum of squares
    SP = 0  # sum of products

    for i in range(len(X)):
        SSx = SSx + (X[i] - meanX) ** 2
        SP = SP + (X[i] - meanX)*(Y[i] - meanY)

    # generates slope and intercept based on standards and baseline samples
    m = SP/SSx
    b = meanY - m * meanX

    SS_res = 0
    SS_t = 0

    for i in range(len(X)):
        SS_res = SS_res + (Y[i] - X[i]*m - b) ** 2
        SS_t = SS_t + (Y[i] - meanY) ** 2
    
    R2 = 1 - SS_res/SS_t    # coefficient of determination

    return m, b, R2

