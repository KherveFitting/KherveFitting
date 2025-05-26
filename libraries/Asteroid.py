import pygame
import math
import random
import sys

# Initialize pygame
pygame.init()

# Constants
SCREEN_WIDTH = 800
SCREEN_HEIGHT = 600
BLACK = (0, 0, 0)
WHITE = (255, 255, 255)
RED = (255, 0, 0)
GREEN = (0, 255, 0)
BLUE = (0, 0, 255)

# Game settings
SHIP_SIZE = 10
SHIP_THRUST = 0.3
SHIP_ROTATION_SPEED = 5
BULLET_SPEED = 10
BULLET_LIFETIME = 60
ASTEROID_SPEEDS = [1, 2, 3]
INITIAL_ASTEROIDS = 5


class Ship:
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self.angle = 0
        self.vel_x = 0
        self.vel_y = 0
        self.radius = SHIP_SIZE

    def update(self):
        # Apply friction
        self.vel_x *= 0.99
        self.vel_y *= 0.99

        # Update position
        self.x += self.vel_x
        self.y += self.vel_y

        # Wrap around screen
        self.x = self.x % SCREEN_WIDTH
        self.y = self.y % SCREEN_HEIGHT

    def thrust(self):
        thrust_x = math.cos(math.radians(self.angle)) * SHIP_THRUST
        thrust_y = math.sin(math.radians(self.angle)) * SHIP_THRUST
        self.vel_x += thrust_x
        self.vel_y += thrust_y

    def rotate_left(self):
        self.angle -= SHIP_ROTATION_SPEED

    def rotate_right(self):
        self.angle += SHIP_ROTATION_SPEED

    def draw(self, screen):
        # Calculate ship points
        cos_a = math.cos(math.radians(self.angle))
        sin_a = math.sin(math.radians(self.angle))

        # Ship vertices (triangle)
        points = [
            (self.x + cos_a * SHIP_SIZE, self.y + sin_a * SHIP_SIZE),
            (self.x + cos_a * (-SHIP_SIZE) + sin_a * (-SHIP_SIZE // 2),
             self.y + sin_a * (-SHIP_SIZE) + cos_a * (SHIP_SIZE // 2)),
            (self.x + cos_a * (-SHIP_SIZE) + sin_a * (SHIP_SIZE // 2),
             self.y + sin_a * (-SHIP_SIZE) + cos_a * (-SHIP_SIZE // 2))
        ]

        pygame.draw.polygon(screen, WHITE, points)


class Bullet:
    def __init__(self, x, y, angle):
        self.x = x
        self.y = y
        self.vel_x = math.cos(math.radians(angle)) * BULLET_SPEED
        self.vel_y = math.sin(math.radians(angle)) * BULLET_SPEED
        self.lifetime = BULLET_LIFETIME
        self.radius = 2

    def update(self):
        self.x += self.vel_x
        self.y += self.vel_y
        self.lifetime -= 1

        # Wrap around screen
        self.x = self.x % SCREEN_WIDTH
        self.y = self.y % SCREEN_HEIGHT

        return self.lifetime > 0

    def draw(self, screen):
        pygame.draw.circle(screen, WHITE, (int(self.x), int(self.y)), self.radius)


class Asteroid:
    def __init__(self, x, y, size):
        self.x = x
        self.y = y
        self.size = size  # 1=small, 2=medium, 3=large
        self.vel_x = random.uniform(-2, 2)
        self.vel_y = random.uniform(-2, 2)
        self.angle = random.uniform(0, 360)
        self.rotation_speed = random.uniform(-5, 5)
        self.radius = size * 15

        # Generate random asteroid shape
        self.points = []
        num_points = 8
        for i in range(num_points):
            angle = (360 / num_points) * i
            distance = self.radius + random.uniform(-5, 5)
            point_x = math.cos(math.radians(angle)) * distance
            point_y = math.sin(math.radians(angle)) * distance
            self.points.append((point_x, point_y))

    def update(self):
        self.x += self.vel_x
        self.y += self.vel_y
        self.angle += self.rotation_speed

        # Wrap around screen
        self.x = self.x % SCREEN_WIDTH
        self.y = self.y % SCREEN_HEIGHT

    def draw(self, screen):
        # Transform points based on position and rotation
        transformed_points = []
        cos_a = math.cos(math.radians(self.angle))
        sin_a = math.sin(math.radians(self.angle))

        for px, py in self.points:
            # Rotate point
            rotated_x = px * cos_a - py * sin_a
            rotated_y = px * sin_a + py * cos_a
            # Translate to position
            final_x = rotated_x + self.x
            final_y = rotated_y + self.y
            transformed_points.append((final_x, final_y))

        pygame.draw.polygon(screen, WHITE, transformed_points, 2)


def check_collision(obj1, obj2):
    distance = math.sqrt((obj1.x - obj2.x) ** 2 + (obj1.y - obj2.y) ** 2)
    return distance < (obj1.radius + obj2.radius)


def create_asteroids(num_asteroids, ship_x, ship_y):
    asteroids = []
    for _ in range(num_asteroids):
        while True:
            x = random.randint(0, SCREEN_WIDTH)
            y = random.randint(0, SCREEN_HEIGHT)
            # Make sure asteroid doesn't spawn too close to ship
            if math.sqrt((x - ship_x) ** 2 + (y - ship_y) ** 2) > 100:
                break

        size = random.choice([1, 2, 3])
        asteroids.append(Asteroid(x, y, size))
    return asteroids


def split_asteroid(asteroid):
    """Split an asteroid into smaller pieces"""
    if asteroid.size <= 1:
        return []

    new_asteroids = []
    new_size = asteroid.size - 1

    for _ in range(2):  # Split into 2 smaller asteroids
        new_asteroid = Asteroid(asteroid.x, asteroid.y, new_size)
        # Give them different velocities
        angle = random.uniform(0, 360)
        speed = random.uniform(1, 3)
        new_asteroid.vel_x = math.cos(math.radians(angle)) * speed
        new_asteroid.vel_y = math.sin(math.radians(angle)) * speed
        new_asteroids.append(new_asteroid)

    return new_asteroids


def main():
    # Reinitialize pygame properly
    if pygame.get_init():
        pygame.quit()
    pygame.init()

    screen = pygame.display.set_mode((SCREEN_WIDTH, SCREEN_HEIGHT))
    pygame.display.set_caption("Meteor smash")
    clock = pygame.time.Clock()

    # Initialize fonts after pygame.init()
    try:
        font = pygame.font.Font(None, 36)
    except pygame.error:
        # Fallback if fonts fail to initialize
        pygame.font.init()
        font = pygame.font.Font(None, 36)

    # Game objects
    ship = Ship(SCREEN_WIDTH // 2, SCREEN_HEIGHT // 2)
    bullets = []
    asteroids = create_asteroids(INITIAL_ASTEROIDS, ship.x, ship.y)

    # Game state
    score = 0
    lives = 3
    game_over = False

    # Input state
    keys_pressed = set()

    running = True
    while running:
        # Handle events
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
            elif event.type == pygame.KEYDOWN:
                keys_pressed.add(event.key)
                if event.key == pygame.K_SPACE and not game_over:
                    # Shoot bullet
                    bullet = Bullet(ship.x, ship.y, ship.angle)
                    bullets.append(bullet)
            elif event.type == pygame.KEYUP:
                keys_pressed.discard(event.key)

        if not game_over:
            # Handle continuous key presses
            if pygame.K_LEFT in keys_pressed or pygame.K_a in keys_pressed:
                ship.rotate_left()
            if pygame.K_RIGHT in keys_pressed or pygame.K_d in keys_pressed:
                ship.rotate_right()
            if pygame.K_UP in keys_pressed or pygame.K_w in keys_pressed:
                ship.thrust()

            # Update game objects
            ship.update()

            # Update bullets
            bullets = [bullet for bullet in bullets if bullet.update()]

            # Update asteroids
            for asteroid in asteroids:
                asteroid.update()

            # Check bullet-asteroid collisions
            for bullet in bullets[:]:
                for asteroid in asteroids[:]:
                    if check_collision(bullet, asteroid):
                        bullets.remove(bullet)
                        asteroids.remove(asteroid)

                        # Add score based on asteroid size
                        score += (4 - asteroid.size) * 20

                        # Split asteroid if it's large enough
                        new_asteroids = split_asteroid(asteroid)
                        asteroids.extend(new_asteroids)
                        break

            # Check ship-asteroid collisions
            for asteroid in asteroids:
                if check_collision(ship, asteroid):
                    lives -= 1
                    # Reset ship position
                    ship.x = SCREEN_WIDTH // 2
                    ship.y = SCREEN_HEIGHT // 2
                    ship.vel_x = 0
                    ship.vel_y = 0
                    if lives <= 0:
                        game_over = True
                    break

            # Check if all asteroids destroyed
            if not asteroids:
                # Create more asteroids for next level
                asteroids = create_asteroids(INITIAL_ASTEROIDS + score // 1000, ship.x, ship.y)

        # Draw everything
        screen.fill(BLACK)

        if not game_over:
            ship.draw(screen)

            for bullet in bullets:
                bullet.draw(screen)

            for asteroid in asteroids:
                asteroid.draw(screen)

        # Draw UI
        score_text = font.render(f"Score: {score}", True, WHITE)
        lives_text = font.render(f"Lives: {lives}", True, WHITE)
        screen.blit(score_text, (10, 10))
        screen.blit(lives_text, (10, 50))

        if game_over:
            game_over_text = font.render("GAME OVER - Press ESC to quit", True, RED)
            text_rect = game_over_text.get_rect(center=(SCREEN_WIDTH // 2, SCREEN_HEIGHT // 2))
            screen.blit(game_over_text, text_rect)

            if pygame.K_ESCAPE in keys_pressed:
                running = False

        # Draw controls
        if not game_over:
            controls = [
                "Arrow Keys / WASD: Move",
                "Space: Shoot"
            ]
            for i, control in enumerate(controls):
                try:
                    control_text = pygame.font.Font(None, 24).render(control, True, WHITE)
                except pygame.error:
                    pygame.font.init()
                    control_text = pygame.font.Font(None, 24).render(control, True, WHITE)
                screen.blit(control_text, (10, SCREEN_HEIGHT - 50 + i * 25))

        pygame.display.flip()
        clock.tick(60)

    pygame.quit()


if __name__ == "__main__":
    main()