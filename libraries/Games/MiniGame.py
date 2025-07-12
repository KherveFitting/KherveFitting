import pygame
import math
import random
from typing import List
import numpy as np

# Initialize Pygame
pygame.init()

# Constants
WIDTH = 400
HEIGHT = 400
FPS = 100

# Colors
WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
GREEN_BG = (144, 238, 144)  # Light green background
MINT_GREEN = (152, 251, 152)  # Lighter mint green, you can adjust these values
PINK_RED = (255, 192, 203)  # Pink-red for electrons

def draw_sphere(screen, pos, radius, color, is_nucleus=False):
    """Draw a sphere with gradient to create 3D effect"""
    # First draw the black outline
    if not is_nucleus:
        pygame.draw.circle(screen, BLACK, pos, radius)  # Add black outline

    # Then draw the inner gradient slightly smaller
    inner_radius = radius - 1 if not is_nucleus else radius
    for i in range(inner_radius, 0, -2):
        # Adjust brightness based on radius
        brightness = int(20 + (radius - i) * 4)  # Make sure brightness is an integer
        # print(f"bright: {brightness}" )
        if is_nucleus:
            # Convert HSV to RGB
            h = color[0] / 360  # Normalize hue to 0-1
            s = color[1] / 100  # Normalize saturation to 0-1
            v = brightness / 255  # Normalize value to 0-1

            # Convert HSV to RGB (simplified conversion)
            c = v * s
            x = c * (1 - abs((h * 6) % 2 - 1))
            m = v - c

            if 0 <= h < 1 / 6:
                r, g, b = c, x, 0
            elif 1 / 6 <= h < 2 / 6:
                r, g, b = x, c, 0
            elif 2 / 6 <= h < 3 / 6:
                r, g, b = 0, c, x
            elif 3 / 6 <= h < 4 / 6:
                r, g, b = 0, x, c
            elif 4 / 6 <= h < 5 / 6:
                r, g, b = x, 0, c
            else:
                r, g, b = c, 0, x

            # Convert to RGB 0-255 range
            current_color = (
                int((r + m) * 255),
                int((g + m) * 255),
                int((b + m) * 255)
            )
        else:
            # Keep electrons blue
            # current_color = (200, 0, 0)#brightness)
            current_color = (199, 12, 120)  # Pink-red for electrons
            # Apply brightness reduction
            # dimmed_color = (
            #     min(255, int(current_color[0] * (brightness / 255))),
            #     min(255, int(current_color[1] * (brightness / 255))),
            #     min(255, int(current_color[2] * (brightness / 255)))
            # )
            # current_color = dimmed_color

        pygame.draw.circle(screen, current_color, pos, i)


class Particle:
    def __init__(self, x, y):
        self.pos = pygame.math.Vector2(x, y)
        self.vel = pygame.math.Vector2(0, 0)
        self.acc = pygame.math.Vector2(0, 0)
        self.tail = []
        self.tail_length = 20
        self.radius = 7  # Radius for the electron



    def move(self, particles, att_force, friction):
        # Update position based on velocity
        self.pos += self.vel
        self.vel += self.acc
        self.vel *= 1 / (friction + 1)

        # Reset acceleration
        self.acc = pygame.math.Vector2(0, 0)

        # Calculate forces from other particles
        for other in particles:
            if other != self:
                d = 1.0 + self.pos.distance_to(other.pos)
                dir = other.pos - self.pos
                try:
                    dir = dir.normalize()
                except ValueError:
                    continue
                dir *= att_force / (d ** 0.95)
                self.acc += dir

        # Update tail
        self.tail.append(pygame.math.Vector2(self.pos))
        if len(self.tail) > self.tail_length:
            self.tail.pop(0)

    def draw(self, screen, color, view_offset):
        # Draw particle tail with view offset and color gradient
        if len(self.tail) > 1:
            screen_points = [(p.x - view_offset.x, p.y - view_offset.y) for p in self.tail]

            # Draw gradient trail segments
            for i in range(len(screen_points) - 1):
                # Calculate color transition
                # Start red at particle end, transition to black
                progress = 1 - (i / (len(screen_points) - 1))  # Reversed progress
                if progress > 0.85:  # First 30% of the trail (nearest to particle)
                    # Transition from black to red
                    red = int(255 * ((progress - 0.85) / 0.15))  # Scale to 0-255
                    trail_color = (red, 0, 0)
                else:
                    trail_color = (0, 0, 0)  # Black

                pygame.draw.line(screen, trail_color, screen_points[i], screen_points[i + 1], 5)

        # Draw particle with view offset and sphere effect
        screen_pos = (int(self.pos.x - view_offset.x), int(self.pos.y - view_offset.y))
        draw_sphere(screen, screen_pos, self.radius, (210, 0, 0))


class ParticleSimulation:
    def __init__(self):
        self.screen = pygame.display.set_mode((WIDTH, HEIGHT))
        pygame.display.set_caption("Electron Game")
        self.clock = pygame.time.Clock()
        self.particles = []
        self.center = pygame.math.Vector2(WIDTH / 2, HEIGHT / 2)
        self.friction = 0.03
        self.att_force = 80
        self.running = True

        # self.back_white = True
        self.back_white = False  # Set to False to use custom background
        self.back_color = GREEN_BG  # Add this line

        self.nucleus_radius = 35
        self.fps = FPS  # Add FPS control

        # New attribute to store last 5 position changes
        self.nucleus_pos_changes = []

        self.prev_nucleus_pos = pygame.math.Vector2(WIDTH / 2, HEIGHT / 2)

        self.stability_time = 0
        self.last_stable_time = 0

        # New attribute for stability plotting
        self.stability_history = []
        self.max_history_length = 100  # Max number of points in the plot

        self.nucleus_color_index = 0
        self.nucleus_colors = [
            (0, 0, 0 ), # black
            (0,0,30), # dark grey
            (120, 100, 100),  # Red
            (120, 100, 100),  # Green
            (210, 100, 100),  # Blue
            (60, 100, 100),  # Yellow
            (270, 100, 100),  # Purple
            (180, 100, 100),  # Cyan
            (300, 100, 100),  # Magenta
        ]

        self.tail_color_index = 0
        self.tail_colors = [
            (210, 250, 200),  # Blue-ish
            (255, 0, 0),      # Red
            (0, 255, 0),      # Green
            (0, 0, 255),      # Blue
            (255, 255, 0),    # Yellow
            (255, 0, 255),    # Magenta
            (0, 255, 255),    # Cyan
        ]

        # Add view offset for camera movement
        self.view_offset = pygame.math.Vector2(0, 0)
        self.camera_speed = 5

        # Load and scale the nucleus image
        # try:
        #     self.nucleus_image = pygame.image.load("nucleus.png").convert_alpha()
        #     # Scale image to match nucleus size (diameter = 2 * radius)
        #     self.nucleus_image = pygame.transform.scale(
        #         self.nucleus_image,
        #         (self.nucleus_radius * 2, self.nucleus_radius * 2)
        #     )
        # except pygame.error:
        #     print("Could not load nucleus image. Using default sphere.")
        #     self.nucleus_image = None

        # Initialize with 3 particles
        for _ in range(3):
            self.particles.append(Particle(random.randint(0, WIDTH),
                                           random.randint(0, HEIGHT)))


    def handle_events(self):
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                self.running = False
                pygame.quit()
                return
            elif event.type == pygame.KEYDOWN:
                if event.key == pygame.K_ESCAPE:
                    self.running = False
                    pygame.quit()
                    return
            elif event.type == pygame.MOUSEBUTTONDOWN:
                if event.button == 1:  # Left click
                    # Add particle at mouse position considering view offset
                    mouse_pos = pygame.mouse.get_pos()
                    real_x = mouse_pos[0] + self.view_offset.x
                    real_y = mouse_pos[1] + self.view_offset.y
                    self.particles.append(Particle(real_x, real_y))
                elif event.button == 3:  # Right click
                    if len(self.particles) > 2:  # Ensure at least 2 particles remain
                        self.particles.pop()

        # Handle keyboard input for camera movement and FPS control
        keys = pygame.key.get_pressed()
        if keys[pygame.K_LEFT]:
            self.view_offset.x -= self.camera_speed
        elif keys[pygame.K_RIGHT]:
            self.view_offset.x += self.camera_speed
        elif keys[pygame.K_UP]:
            self.view_offset.y -= self.camera_speed
        elif keys[pygame.K_DOWN]:
            self.view_offset.y += self.camera_speed
        elif keys[pygame.K_EQUALS] or keys[pygame.K_KP_EQUALS]:
            self.fps = min(200, self.fps + 1)  # Cap at 200 FPS
        elif keys[pygame.K_MINUS] or keys[pygame.K_KP_MINUS]:
            self.fps = max(10, self.fps - 1)   # Minimum 10 FPS
        elif keys[pygame.K_c]:  # Add 'c' key for color change
            # print("Change colors")
            self.tail_color_index = (self.tail_color_index + 1) % len(self.tail_colors)
            self.nucleus_color_index = (self.nucleus_color_index + 1) % len(self.nucleus_colors)

    def update(self):
        # Update particle positions
        for p in self.particles:
            p.move(self.particles, self.att_force, self.friction)

        # Update center position
        if self.particles:
            new_center = pygame.math.Vector2(0, 0)
            for p in self.particles:
                new_center += p.pos
            self.center = new_center / len(self.particles)

        # Update view offset to follow the particles and nucleus
        target_offset = pygame.math.Vector2(
            self.center.x - WIDTH // 2,
            self.center.y - HEIGHT // 2
        )
        self.view_offset = self.view_offset.lerp(target_offset, 0.05)

        # Calculate stability based on change in nucleus position
        nucleus_pos_change = self.center.distance_to(self.prev_nucleus_pos)

        # Update the list of position changes, keeping only the last 5 values
        self.nucleus_pos_changes.append(nucleus_pos_change)
        if len(self.nucleus_pos_changes) > 100:
            self.nucleus_pos_changes.pop(0)

        # Calculate the average position change over the last 100 frames
        avg_nucleus_pos_change = sum(self.nucleus_pos_changes) / len(self.nucleus_pos_changes)

        # Add the stability value to the history for plotting
        self.stability_history.append(avg_nucleus_pos_change)
        if len(self.stability_history) > self.max_history_length:
            self.stability_history.pop(0)

        # Adjust FPS based on stability
        if avg_nucleus_pos_change > 1:
            self.fps = 150
        elif avg_nucleus_pos_change < 0.5:
            self.fps = 60

        # Use this average for stability calculations
        self.stability_value = avg_nucleus_pos_change
        self.stability = self.calculate_stability(avg_nucleus_pos_change)

        # Store current center as the previous position for the next frame
        self.prev_nucleus_pos = self.center.copy()

        # Update stability time and last stable time
        if self.stability == "Very Stable" or self.stability == "Stable":
            self.stability_time += 1 / self.fps  # Increment stability time by the frame duration
        else:
            if self.stability_time > 0:  # Only update last_stable_time if we were previously stable
                self.last_stable_time = self.stability_time
            self.stability_time = 0
    def calculate_stability(self, nucleus_pos_change):
        if nucleus_pos_change <= 0.3:
            return "Very Stable"
        elif nucleus_pos_change <= 0.5:
            return "Stable"
        else:
            return "Unstable"

    def draw(self):
        # Clear screen
        # self.screen.fill(WHITE if self.back_white else BLACK)
        self.screen.fill(self.back_color)


        # Draw nucleus with view offset and sphere effect
        nucleus_screen_pos = (int(self.center.x - self.view_offset.x),
                              int(self.center.y - self.view_offset.y))

        # Get current nucleus color
        current_nucleus_color = self.nucleus_colors[self.nucleus_color_index]

        self.nucleus_image= False
        if self.nucleus_image:
            # Calculate position to center the image
            image_pos = (nucleus_screen_pos[0] - self.nucleus_radius,
                         nucleus_screen_pos[1] - self.nucleus_radius)
            self.screen.blit(self.nucleus_image, image_pos)
        else:
            # Fallback to original sphere drawing if image loading failed
            draw_sphere(screen=self.screen,
                        pos=nucleus_screen_pos,
                        radius=self.nucleus_radius,
                        color=current_nucleus_color,
                        is_nucleus=True)
        # Draw particles with view offset
        particle_color = BLACK if self.back_white else WHITE
        for p in self.particles:
            p.draw(self.screen, particle_color, self.view_offset)

        # Draw help text and shortcuts
        font = pygame.font.SysFont('Arial', 16)
        # First define GREY at the top of your file with other colors
        GREY = (128, 128, 128)  # Medium grey - you can adjust these values for lighter/darker grey

        # Then in your draw method, modify the text_color line to:
        text_color = GREY if self.back_white else GREY

        # Create list of help texts
        help_texts = [
            f"Electrons: {len(self.particles)}",
            f"Stability: {self.stability}",
            f"Stability: {self.stability_value:.2f}",
            f"Last Stable Time: {self.last_stable_time:.2f}s",
            "",
            "Controls:",
            f"FPS: {self.fps}",
            "↑↓←→: Move view",
            # "+/-: Change speed",
            "c: Change color",
            "Left click: Add electron",
            "Right click: Remove electron"
        ]

        # Draw all help texts
        for i, text in enumerate(help_texts):
            text_surface = font.render(text, True, text_color)
            self.screen.blit(text_surface, (WIDTH - 390, 10 + i * 20))

        # Draw the stability plot at the bottom
        self.draw_stability_plot()



        pygame.display.flip()

    def draw_stability_plot(self):
        plot_height = 50
        plot_width = WIDTH
        plot_x = 0
        plot_y = HEIGHT - plot_height - 10

        # Draw plot background
        # pygame.draw.rect(self.screen, (255, 255, 255), (plot_x, plot_y, plot_width, plot_height))

        if len(self.stability_history) > 1:
            # Use actual min/max for scaling
            max_stability_value = max(self.stability_history)
            min_stability_value = min(self.stability_history)
            value_range = max_stability_value - min_stability_value or 1  # Avoid division by zero

            # Only plot every 300th point
            plot_interval = 300
            filtered_history = self.stability_history[::plot_interval]

            # Make sure we have at least 2 points
            if len(filtered_history) < 2:
                filtered_history = self.stability_history  # Use all points if we have too few

            # Scale points to fit plot height
            scaled_points = [
                (plot_x + i * (plot_width / (len(filtered_history) - 1 or 1)),
                 plot_y + plot_height - ((value - min_stability_value) / value_range * plot_height))
                for i, value in enumerate(filtered_history)
            ]

            # Only draw lines if we have at least 2 points
            if len(scaled_points) >= 2:
                pygame.draw.lines(self.screen, (240, 240, 240), False, scaled_points, 2)

            # Draw Y-axis labels and grid lines
            font = pygame.font.SysFont('Arial', 12)
            num_y_intervals = 2  # Number of intervals to show

            for i in range(num_y_intervals + 1):
                # Calculate value for this interval
                value = min_stability_value + (i * value_range / num_y_intervals)
                # Calculate y position
                y_pos = plot_y + plot_height - (i * plot_height / num_y_intervals)



                # Draw y-axis label
                label = font.render(f"{value:.3f}", True, (240, 240, 240))
                self.screen.blit(label, (plot_x + 5, y_pos - 6))

    def run(self):
        try:
            while self.running:
                self.handle_events()
                if not self.running:
                    break
                self.update()
                self.draw()
                self.clock.tick(self.fps)
        except:
            pass
        finally:
            pygame.quit()


if __name__ == "__main__":
    sim = ParticleSimulation()
    sim.run()
    pygame.quit()