import pygame
import random
import sys
import math

# Initialize Pygame
pygame.init()

# Constants
WINDOW_WIDTH = 400
WINDOW_HEIGHT = 600
BIRD_WIDTH = 40
BIRD_HEIGHT = 30
PIPE_WIDTH = 80
PIPE_GAP = 180
GRAVITY = 0.4
JUMP_STRENGTH = -7

# Difficulty settings
DIFFICULTIES = {
    'easy': {'pipe_speed': 2, 'pipe_gap': 220, 'enemy_spawn_rate': 0.003},
    'medium': {'pipe_speed': 3, 'pipe_gap': 180, 'enemy_spawn_rate': 0.005},
    'hard': {'pipe_speed': 4, 'pipe_gap': 140, 'enemy_spawn_rate': 0.008}
}

# Colors
SKY_BLUE = (135, 206, 235)
CLOUD_WHITE = (255, 255, 255)
PIPE_GREEN = (34, 139, 34)
PIPE_DARK_GREEN = (0, 100, 0)
BIRD_YELLOW = (255, 215, 0)
BIRD_ORANGE = (255, 140, 0)
WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
GROUND_BROWN = (139, 69, 19)
GRASS_GREEN = (34, 139, 34)
RED = (255, 0, 0)


class Sun:
    def __init__(self):
        self.time = 0  # 0 to 1440 (24 hours in minutes)
        self.speed = 0.5  # Speed of day cycle

    def update(self):
        self.time = (self.time + self.speed) % 1440

    def get_position(self):
        # Sun rises at 360 (6am) and sets at 1080 (6pm)
        if 360 <= self.time <= 1080:  # Daytime
            # Calculate arc position - wider arc going to ground
            progress = (self.time - 360) / 720  # 0 to 1
            angle = math.pi * progress  # 0 to pi
            x = WINDOW_WIDTH * progress
            y = (WINDOW_HEIGHT - 100) - math.sin(angle) * (WINDOW_HEIGHT - 100)
            return (int(x), int(y))
        return None

    def get_moon_position(self):
        # Moon is visible at night
        if self.time < 360 or self.time > 1080:  # Nighttime
            if self.time > 1080:
                progress = (self.time - 1080) / 360  # 0 to 1
            else:
                progress = (self.time + 360) / 360  # 0 to 1
            angle = math.pi * progress
            x = WINDOW_WIDTH * progress
            y = (WINDOW_HEIGHT - 100) - math.sin(angle) * (WINDOW_HEIGHT - 100)
            return (int(x), int(y))
        return None

    def get_sky_color(self):
        # Dawn: 300-420 (5am-7am)
        # Day: 420-1020 (7am-5pm)
        # Dusk: 1020-1140 (5pm-7pm)
        # Night: 1140-300 (7pm-5am)

        if 420 <= self.time <= 1020:  # Day
            return (135, 206, 235)
        elif 1140 <= self.time or self.time <= 300:  # Night
            return (25, 25, 60)
        elif 300 < self.time < 420:  # Dawn - orangey sunrise
            progress = (self.time - 300) / 120
            r = int(255 - (255 - 135) * progress)
            g = int(140 + (206 - 140) * progress)
            b = int(60 + (235 - 60) * progress)
            return (r, g, b)
        else:  # Dusk (1020-1140) - orangey sunset
            progress = (self.time - 1020) / 120
            r = int(135 + (255 - 135) * progress)
            g = int(206 - (206 - 140) * progress)
            b = int(235 - (235 - 60) * progress)
            return (r, g, b)

    def draw(self, screen):
        sun_pos = self.get_position()
        moon_pos = self.get_moon_position()

        if sun_pos:
            # Draw sun with glow
            pygame.draw.circle(screen, (255, 255, 150), sun_pos, 35)
            pygame.draw.circle(screen, (255, 255, 0), sun_pos, 30)
            pygame.draw.circle(screen, (255, 220, 0), sun_pos, 25)

        if moon_pos:
            # Draw moon
            pygame.draw.circle(screen, (240, 240, 255), moon_pos, 25)
            pygame.draw.circle(screen, (220, 220, 240), moon_pos, 23)
            # Moon craters
            pygame.draw.circle(screen, (200, 200, 220), (moon_pos[0] - 5, moon_pos[1] - 3), 4)
            pygame.draw.circle(screen, (200, 200, 220), (moon_pos[0] + 7, moon_pos[1] + 5), 3)


class Bullet:
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self.speed = 8
        self.rect = pygame.Rect(x, y, 12, 8)

    def update(self):
        self.x += self.speed
        self.rect.x = self.x

    def draw(self, screen):
        pygame.draw.ellipse(screen, (255, 255, 0), self.rect)
        pygame.draw.ellipse(screen, (255, 200, 0), self.rect, 2)

    def is_off_screen(self):
        return self.x > WINDOW_WIDTH


class EnemyBird:
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self.speed = 2
        self.rect = pygame.Rect(x, y, BIRD_WIDTH, BIRD_HEIGHT)
        self.flap_animation = 0
        self.vertical_speed = 0

        # Random bird colors
        colors = [
            (255, 0, 0),  # Red
            (138, 43, 226),  # Purple
            (0, 191, 255),  # Blue
            (255, 20, 147),  # Pink
            (255, 140, 0),  # Orange
            (50, 205, 50)  # Green
        ]
        self.color = random.choice(colors)
        self.dark_color = tuple(max(0, c - 100) for c in self.color)
        self.light_color = tuple(min(255, c + 50) for c in self.color)

    def update(self, main_bird):
        self.x -= self.speed
        self.rect.x = self.x
        self.flap_animation = (self.flap_animation + 1) % 20

        # Move opposite to main bird's velocity
        if main_bird.velocity > 0:  # Main bird falling
            self.vertical_speed = -1.5  # Enemy bird goes up
        elif main_bird.velocity < 0:  # Main bird rising
            self.vertical_speed = 1.5  # Enemy bird goes down
        else:
            self.vertical_speed *= 0.9  # Slow down when main bird is stable

        self.y += self.vertical_speed

        # Keep enemy birds within bounds
        if self.y < 50:
            self.y = 50
            self.vertical_speed = 0
        elif self.y > WINDOW_HEIGHT - 150:
            self.y = WINDOW_HEIGHT - 150
            self.vertical_speed = 0

        self.rect.y = self.y

    def draw(self, screen):
        # Draw enemy bird with random color
        bird_surface = pygame.Surface((BIRD_WIDTH, BIRD_HEIGHT), pygame.SRCALPHA)

        # Draw bird body (ellipse)
        pygame.draw.ellipse(bird_surface, self.color, (0, 0, BIRD_WIDTH, BIRD_HEIGHT))
        pygame.draw.ellipse(bird_surface, self.dark_color, (0, 0, BIRD_WIDTH, BIRD_HEIGHT), 3)

        # Draw wing (animated)
        wing_offset = 5 if self.flap_animation > 10 else 0
        pygame.draw.ellipse(bird_surface, self.light_color,
                            (8, 8 - wing_offset, BIRD_WIDTH // 2, BIRD_HEIGHT // 2))

        # Draw eye
        pygame.draw.circle(bird_surface, WHITE, (10, BIRD_HEIGHT // 2 - 3), 6)
        pygame.draw.circle(bird_surface, BLACK, (12, BIRD_HEIGHT // 2 - 3), 3)

        # Draw beak (pointing left)
        beak_points = [(0, BIRD_HEIGHT // 2), (-8, BIRD_HEIGHT // 2 - 2),
                       (-8, BIRD_HEIGHT // 2 + 2)]
        pygame.draw.polygon(bird_surface, self.dark_color, beak_points)

        screen.blit(bird_surface, (self.x, self.y))

    def is_off_screen(self):
        return self.x + BIRD_WIDTH < 0


class Bee:
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self.speed = 3
        self.rect = pygame.Rect(x, y, 25, 20)
        self.wing_animation = 0
        self.hover_offset = 0
        self.hover_speed = 0.2

    def update(self):
        self.x -= self.speed
        self.rect.x = self.x
        self.wing_animation = (self.wing_animation + 1) % 10

        # Hovering motion
        self.hover_offset += self.hover_speed
        self.y += math.sin(self.hover_offset) * 0.5
        self.rect.y = int(self.y)

    def draw(self, screen):
        bee_surface = pygame.Surface((25, 20), pygame.SRCALPHA)

        # Draw body (yellow with black stripes)
        pygame.draw.ellipse(bee_surface, (255, 215, 0), (5, 5, 15, 12))

        # Black stripes
        pygame.draw.line(bee_surface, BLACK, (8, 5), (8, 17), 2)
        pygame.draw.line(bee_surface, BLACK, (13, 5), (13, 17), 2)
        pygame.draw.line(bee_surface, BLACK, (17, 5), (17, 17), 2)

        # Wings (animated)
        wing_offset = 3 if self.wing_animation < 5 else 0
        pygame.draw.ellipse(bee_surface, (200, 230, 255), (3, 2 - wing_offset, 8, 6))
        pygame.draw.ellipse(bee_surface, (200, 230, 255), (14, 2 - wing_offset, 8, 6))

        # Head
        pygame.draw.circle(bee_surface, BLACK, (5, 10), 4)

        # Eyes
        pygame.draw.circle(bee_surface, WHITE, (4, 9), 2)
        pygame.draw.circle(bee_surface, BLACK, (4, 9), 1)

        # Stinger
        pygame.draw.polygon(bee_surface, BLACK, [(20, 11), (24, 10), (24, 12)])

        screen.blit(bee_surface, (int(self.x), int(self.y)))

    def is_off_screen(self):
        return self.x + 25 < 0


class Particle:
    def __init__(self, x, y, color, velocity):
        self.x = x
        self.y = y
        self.color = color
        self.velocity = velocity
        self.life = 30
        self.max_life = 30

    def update(self):
        self.x += self.velocity[0]
        self.y += self.velocity[1]
        self.life -= 1

    def draw(self, screen):
        alpha = int(255 * (self.life / self.max_life))
        size = int(3 * (self.life / self.max_life))
        if size > 0:
            pygame.draw.circle(screen, self.color, (int(self.x), int(self.y)), size)


class Cloud:
    def __init__(self, x, y, size):
        self.x = x
        self.y = y
        self.size = size
        self.speed = random.uniform(0.2, 0.5)

    def update(self):
        self.x -= self.speed
        if self.x + self.size < 0:
            self.x = WINDOW_WIDTH + random.randint(50, 200)
            self.y = random.randint(50, 200)

    def draw(self, screen):
        pygame.draw.circle(screen, CLOUD_WHITE, (int(self.x), int(self.y)), self.size)
        pygame.draw.circle(screen, CLOUD_WHITE, (int(self.x + self.size // 2), int(self.y)), self.size // 2)
        pygame.draw.circle(screen, CLOUD_WHITE, (int(self.x - self.size // 2), int(self.y)), self.size // 2)
        pygame.draw.circle(screen, CLOUD_WHITE, (int(self.x), int(self.y - self.size // 3)), self.size // 2)


class Flower:
    def __init__(self, x):
        self.x = x
        self.y = WINDOW_HEIGHT - 100  # Ground level
        self.petal_color = random.choice([
            (255, 105, 180),  # Pink
            (255, 0, 0),  # Red
            (255, 255, 0),  # Yellow
            (138, 43, 226),  # Purple
            (255, 165, 0)  # Orange
        ])
        self.center_color = (255, 215, 0)  # Yellow center
        self.size = random.randint(8, 12)

    def update(self, speed):
        self.x -= speed
        if self.x < -20:
            self.x = WINDOW_WIDTH + random.randint(0, 100)

    def draw(self, screen):
        # Draw stem
        pygame.draw.line(screen, GRASS_GREEN,
                         (int(self.x), int(self.y)),
                         (int(self.x), int(self.y) + 20), 2)

        # Draw petals (5 petals in a circle)
        for angle in range(0, 360, 72):
            rad = math.radians(angle)
            petal_x = self.x + math.cos(rad) * (self.size // 2)
            petal_y = self.y + math.sin(rad) * (self.size // 2)
            pygame.draw.circle(screen, self.petal_color,
                               (int(petal_x), int(petal_y)), self.size // 2)

        # Draw center
        pygame.draw.circle(screen, self.center_color,
                           (int(self.x), int(self.y)), self.size // 3)


class Stone:
    def __init__(self, x):
        self.x = x
        self.y = WINDOW_HEIGHT - 100
        self.size = random.randint(15, 25)
        self.color = random.choice([
            (128, 128, 128),
            (105, 105, 105),
            (169, 169, 169)
        ])
        # Generate fixed shape once
        self.points = []
        num_points = 8
        for i in range(num_points):
            angle = (2 * math.pi * i) / num_points
            radius = self.size // 2 + random.randint(-3, 3)
            px = math.cos(angle) * radius
            py = math.sin(angle) * radius
            self.points.append((px, py))

    def update(self, speed):
        self.x -= speed
        if self.x < -self.size:
            self.x = WINDOW_WIDTH + random.randint(0, 100)

    def draw(self, screen):
        # Draw stone with fixed shape
        points = [(int(self.x + p[0]), int(self.y + p[1])) for p in self.points]
        pygame.draw.polygon(screen, self.color, points)
        pygame.draw.polygon(screen, (80, 80, 80), points, 2)


class Bird:
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self.velocity = 0
        self.rect = pygame.Rect(x, y, BIRD_WIDTH, BIRD_HEIGHT)
        self.angle = 0
        self.flap_animation = 0

    def jump(self):
        self.velocity = JUMP_STRENGTH
        self.flap_animation = 10
        return [(self.x + random.randint(-5, 5), self.y + random.randint(-5, 5),
                 BIRD_ORANGE, (random.uniform(-2, 2), random.uniform(-3, -1))) for _ in range(3)]

    def shoot(self):
        return Bullet(self.x + BIRD_WIDTH, self.y + BIRD_HEIGHT // 2)

    def update(self):
        self.velocity += GRAVITY
        self.y += self.velocity
        self.rect.y = self.y

        self.angle = max(-30, min(30, self.velocity * 3))

        if self.flap_animation > 0:
            self.flap_animation -= 1

        if self.y < 0:
            self.y = 0
            self.velocity = 0

    def draw(self, screen):
        bird_surface = pygame.Surface((BIRD_WIDTH, BIRD_HEIGHT), pygame.SRCALPHA)

        pygame.draw.ellipse(bird_surface, BIRD_YELLOW, (0, 0, BIRD_WIDTH, BIRD_HEIGHT))
        pygame.draw.ellipse(bird_surface, BIRD_ORANGE, (0, 0, BIRD_WIDTH, BIRD_HEIGHT), 3)

        wing_offset = 5 if self.flap_animation > 5 else 0
        pygame.draw.ellipse(bird_surface, BIRD_ORANGE,
                            (8, 8 - wing_offset, BIRD_WIDTH // 2, BIRD_HEIGHT // 2))

        pygame.draw.circle(bird_surface, WHITE, (BIRD_WIDTH - 10, BIRD_HEIGHT // 2 - 3), 6)
        pygame.draw.circle(bird_surface, BLACK, (BIRD_WIDTH - 8, BIRD_HEIGHT // 2 - 3), 3)

        beak_points = [(BIRD_WIDTH, BIRD_HEIGHT // 2), (BIRD_WIDTH + 8, BIRD_HEIGHT // 2 - 2),
                       (BIRD_WIDTH + 8, BIRD_HEIGHT // 2 + 2)]
        pygame.draw.polygon(bird_surface, BIRD_ORANGE, beak_points)

        rotated_bird = pygame.transform.rotate(bird_surface, -self.angle)
        rotated_rect = rotated_bird.get_rect(center=(self.x + BIRD_WIDTH // 2, self.y + BIRD_HEIGHT // 2))
        screen.blit(rotated_bird, rotated_rect)


class Pipe:
    def __init__(self, x, pipe_speed, pipe_gap):
        self.x = x
        self.speed = pipe_speed
        self.gap_y = random.randint(150, WINDOW_HEIGHT - pipe_gap - 150)
        self.gap_size = pipe_gap
        self.passed = False
        self.show_text = random.random() < 0.25  # 25% chance to show text
        self.has_stone_top = random.random() < 0.1  # 10% chance for stone on top pipe
        self.has_stone_bottom = random.random() < 0.1  # 10% chance for stone on bottom pipe

        # Create rectangles for collision detection
        self.top_rect = pygame.Rect(x, 0, PIPE_WIDTH, self.gap_y)
        self.bottom_rect = pygame.Rect(x, self.gap_y + pipe_gap, PIPE_WIDTH,
                                       WINDOW_HEIGHT - (self.gap_y + pipe_gap) - 100)

    def update(self):
        self.x -= self.speed
        self.top_rect.x = self.x
        self.bottom_rect.x = self.x

    def draw(self, screen):
        # Draw top pipe
        pygame.draw.rect(screen, PIPE_GREEN, self.top_rect)
        pygame.draw.rect(screen, PIPE_DARK_GREEN, self.top_rect, 3)
        pygame.draw.rect(screen, PIPE_DARK_GREEN, (self.x, self.gap_y - 20, PIPE_WIDTH, 20))

        # Draw bottom pipe
        pygame.draw.rect(screen, PIPE_GREEN, self.bottom_rect)
        pygame.draw.rect(screen, PIPE_DARK_GREEN, self.bottom_rect, 3)
        pygame.draw.rect(screen, PIPE_DARK_GREEN, (self.x, self.gap_y + self.gap_size, PIPE_WIDTH, 20))

        # Draw highlights for 3D effect
        pygame.draw.line(screen, (50, 205, 50), (self.x + 5, 0), (self.x + 5, self.gap_y), 2)
        pygame.draw.line(screen, (50, 205, 50), (self.x + 5, self.gap_y + self.gap_size),
                         (self.x + 5, WINDOW_HEIGHT - 100), 2)

        # Draw stone on top of top pipe if present
        if self.has_stone_top:
            stone_x = self.x + PIPE_WIDTH // 2
            stone_y = self.gap_y - 35
            stone_size = 20
            stone_color = (128, 128, 128)

            # Fixed stone shape
            points = []
            num_points = 6
            base_offsets = [3, -2, 2, -3, 1, -1]  # Fixed offsets
            for i in range(num_points):
                angle = (2 * math.pi * i) / num_points
                radius = stone_size // 2 + base_offsets[i]
                px = stone_x + math.cos(angle) * radius
                py = stone_y + math.sin(angle) * radius
                points.append((int(px), int(py)))

            pygame.draw.polygon(screen, stone_color, points)
            pygame.draw.polygon(screen, (80, 80, 80), points, 2)

        # Draw stone on top of bottom pipe if present
        if self.has_stone_bottom:
            stone_x = self.x + PIPE_WIDTH // 2
            stone_y = self.gap_y + self.gap_size + 5
            stone_size = 20
            stone_color = (128, 128, 128)

            # Fixed stone shape
            points = []
            num_points = 6
            base_offsets = [3, -2, 2, -3, 1, -1]  # Fixed offsets
            for i in range(num_points):
                angle = (2 * math.pi * i) / num_points
                radius = stone_size // 2 + base_offsets[i]
                px = stone_x + math.cos(angle) * radius
                py = stone_y + math.sin(angle) * radius
                points.append((int(px), int(py)))

            pygame.draw.polygon(screen, stone_color, points)
            pygame.draw.polygon(screen, (80, 80, 80), points, 2)

        if self.show_text:
            font = pygame.font.Font(None, 20)
            text = font.render("XPS", True, WHITE)
            text_rect = text.get_rect(center=(self.x + PIPE_WIDTH // 2, self.gap_y + self.gap_size // 2))
            pygame.draw.rect(screen, PIPE_DARK_GREEN, text_rect.inflate(10, 10))
            screen.blit(text, text_rect)

    def collides_with(self, bird):
        return self.top_rect.colliderect(bird.rect) or self.bottom_rect.colliderect(bird.rect)

    def is_off_screen(self):
        return self.x + PIPE_WIDTH < 0


def draw_ground(screen, offset):
    # Draw grass
    pygame.draw.rect(screen, GRASS_GREEN, (0, WINDOW_HEIGHT - 100, WINDOW_WIDTH, 20))

    # Draw dirt/ground
    for y in range(WINDOW_HEIGHT - 80, WINDOW_HEIGHT):
        darkness = (y - (WINDOW_HEIGHT - 80)) / 80
        r = int(GROUND_BROWN[0] * (1 - darkness * 0.3))
        g = int(GROUND_BROWN[1] * (1 - darkness * 0.3))
        b = int(GROUND_BROWN[2] * (1 - darkness * 0.3))
        pygame.draw.line(screen, (r, g, b), (0, y), (WINDOW_WIDTH, y))

    # Draw grass blades
    for i in range(0, WINDOW_WIDTH, 20):
        x = (i + offset) % WINDOW_WIDTH
        pygame.draw.line(screen, (0, 128, 0), (x, WINDOW_HEIGHT - 100), (x - 2, WINDOW_HEIGHT - 105), 2)
        pygame.draw.line(screen, (0, 128, 0), (x + 5, WINDOW_HEIGHT - 100), (x + 3, WINDOW_HEIGHT - 108), 2)
        pygame.draw.line(screen, (0, 128, 0), (x + 10, WINDOW_HEIGHT - 100), (x + 12, WINDOW_HEIGHT - 106), 2)


def show_menu(screen, font, big_font):
    menu_running = True
    selected = 1  # 0=easy, 1=medium, 2=hard

    while menu_running:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                return None
            elif event.type == pygame.KEYDOWN:
                if event.key == pygame.K_UP:
                    selected = (selected - 1) % 3
                elif event.key == pygame.K_DOWN:
                    selected = (selected + 1) % 3
                elif event.key == pygame.K_RETURN or event.key == pygame.K_SPACE:
                    return ['easy', 'medium', 'hard'][selected]
                elif event.key == pygame.K_q:
                    return None

        screen.fill(SKY_BLUE)

        title = big_font.render("Khervey the Flappy Bird", True, WHITE)
        title_outline = big_font.render("Khervey the Flappy Bird", True, BLACK)
        screen.blit(title_outline, (WINDOW_WIDTH // 2 - 192, 102))
        screen.blit(title, (WINDOW_WIDTH // 2 - 190, 100))

        subtitle = font.render("Select Difficulty", True, WHITE)
        screen.blit(subtitle, (WINDOW_WIDTH // 2 - 80, 180))

        difficulties = ['Easy', 'Medium', 'Hard']
        y_positions = [250, 300, 350]

        for i, diff in enumerate(difficulties):
            color = BIRD_YELLOW if i == selected else WHITE
            outline_color = BIRD_ORANGE if i == selected else BLACK
            text = font.render(diff, True, color)
            text_outline = font.render(diff, True, outline_color)

            screen.blit(text_outline, (WINDOW_WIDTH // 2 - 32, y_positions[i] + 2))
            screen.blit(text, (WINDOW_WIDTH // 2 - 30, y_positions[i]))

            if i == selected:
                pygame.draw.rect(screen, BIRD_ORANGE,
                                 (WINDOW_WIDTH // 2 - 50, y_positions[i] - 5, 100, 35), 3)

        controls = font.render("Controls: UP=Jump, SPACE=Shoot", True, WHITE)
        screen.blit(controls, (WINDOW_WIDTH // 2 - 150, 450))

        pygame.display.flip()

    return None


def main():
    # Force complete pygame shutdown and restart
    try:
        pygame.quit()
    except:
        pass

    # Reinitialize pygame completely
    pygame.init()
    pygame.font.init()  # Explicitly initialize fonts
    pygame.mixer.init()  # Initialize mixer too

    # Verify initialization
    if not pygame.get_init():
        print("Failed to initialize pygame")
        return

    if not pygame.font.get_init():
        print("Failed to initialize pygame fonts")
        pygame.font.init()

    screen = pygame.display.set_mode((WINDOW_WIDTH, WINDOW_HEIGHT))
    pygame.display.set_caption("Khervey the flappy bird")
    clock = pygame.time.Clock()

    # Initialize fonts with error handling
    font = None
    big_font = None
    for attempt in range(3):  # Try 3 times
        try:
            font = pygame.font.Font(None, 36)
            big_font = pygame.font.Font(None, 48)
            break
        except pygame.error as e:
            print(f"Font initialization attempt {attempt + 1} failed: {e}")
            pygame.font.quit()
            pygame.font.init()
            if attempt == 2:  # Last attempt
                print("Failed to initialize fonts after 3 attempts")
                pygame.quit()
                return

    # Show menu and get difficulty
    difficulty = show_menu(screen, font, big_font)
    if difficulty is None:
        pygame.quit()
        return

    # Set difficulty parameters
    settings = DIFFICULTIES[difficulty]
    pipe_speed = settings['pipe_speed']
    pipe_gap = settings['pipe_gap']
    enemy_spawn_rate = settings['enemy_spawn_rate']

    bird = Bird(50, WINDOW_HEIGHT // 2)
    pipes = []
    clouds = [Cloud(random.randint(0, WINDOW_WIDTH), random.randint(50, 200),
                    random.randint(20, 40)) for _ in range(5)]
    flowers = [Flower(random.randint(0, WINDOW_WIDTH)) for _ in range(10)]
    stones = [Stone(random.randint(0, WINDOW_WIDTH)) for _ in range(8)]
    particles = []
    bullets = []
    enemy_birds = []
    bees = []
    score = 0
    stage = 1
    stage_thresholds = [20, 40, 70, 100, 140, 190, 250, 320, 400, 500]
    game_over = False
    ground_offset = 0
    shoot_cooldown = 0
    shoot_delay = 10  # Frames between shots
    sun = Sun()

    pipes.append(Pipe(WINDOW_WIDTH, pipe_speed, pipe_gap))

    running = True
    while running:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
            elif event.type == pygame.KEYDOWN:
                if event.key == pygame.K_q:  # Add Q to quit
                    running = False
                elif event.key == pygame.K_UP and not game_over:
                    new_particles = bird.jump()
                    particles.extend([Particle(p[0], p[1], p[2], p[3]) for p in new_particles])
                elif (event.key == pygame.K_DOWN or event.key == pygame.K_r) and game_over:
                    # Restart with same difficulty
                    bird = Bird(50, WINDOW_HEIGHT // 2)
                    pipes = [Pipe(WINDOW_WIDTH, pipe_speed, pipe_gap)]
                    flowers = [Flower(random.randint(0, WINDOW_WIDTH)) for _ in range(10)]
                    stones = [Stone(random.randint(0, WINDOW_WIDTH)) for _ in range(8)]
                    particles = []
                    bullets = []
                    enemy_birds = []
                    bees = []
                    score = 0
                    stage = 1
                    game_over = False
                    shoot_cooldown = 0

        if not game_over:
            bird.update()
            ground_offset = (ground_offset - pipe_speed) % 20
            sun.update()

            # Continuous shooting when space is held
            keys = pygame.key.get_pressed()
            if keys[pygame.K_SPACE] and shoot_cooldown <= 0:
                bullets.append(bird.shoot())
                shoot_cooldown = shoot_delay

            if shoot_cooldown > 0:
                shoot_cooldown -= 1

            if bird.y + BIRD_HEIGHT >= WINDOW_HEIGHT - 100:
                game_over = True

            # Update pipes
            for pipe in pipes[:]:
                pipe.update()

                if pipe.collides_with(bird):
                    game_over = True
                    for _ in range(10):
                        particles.append(Particle(bird.x + BIRD_WIDTH // 2, bird.y + BIRD_HEIGHT // 2,
                                                  (255, random.randint(0, 100), 0),
                                                  (random.uniform(-4, 4), random.uniform(-4, 4))))

                if not pipe.passed and pipe.x + PIPE_WIDTH < bird.x:
                    pipe.passed = True
                    score += 1

                    # Check for stage advancement
                    if stage <= len(stage_thresholds) and score >= stage_thresholds[stage - 1]:
                        stage += 1
                        # Extra particles for stage up
                        for _ in range(15):
                            particles.append(Particle(bird.x, bird.y,
                                                      (255, 215, 0),
                                                      (random.uniform(-3, 3), random.uniform(-4, 4))))
                    else:
                        for _ in range(5):
                            particles.append(Particle(bird.x, bird.y,
                                                      (255, 255, 0),
                                                      (random.uniform(-2, 2), random.uniform(-2, 2))))

                if pipe.is_off_screen():
                    pipes.remove(pipe)

            if len(pipes) == 0 or pipes[-1].x < WINDOW_WIDTH - 200:
                pipes.append(Pipe(WINDOW_WIDTH, pipe_speed, pipe_gap))

            # Spawn enemy birds and bees
            if random.random() < enemy_spawn_rate:
                enemy_y = random.randint(50, WINDOW_HEIGHT - 150)
                enemy_birds.append(EnemyBird(WINDOW_WIDTH, enemy_y))

            if random.random() < enemy_spawn_rate * 1.5:  # Bees spawn more frequently
                bee_y = random.randint(50, WINDOW_HEIGHT - 150)
                bees.append(Bee(WINDOW_WIDTH, bee_y))

            # Update enemy birds
            for enemy in enemy_birds[:]:
                enemy.update(bird)  # Pass main bird to enemy for opposite movement
                if enemy.is_off_screen():
                    enemy_birds.remove(enemy)
                elif enemy.rect.colliderect(bird.rect):
                    game_over = True

            # Update bees
            for bee in bees[:]:
                bee.update()
                if bee.is_off_screen():
                    bees.remove(bee)
                elif bee.rect.colliderect(bird.rect):
                    game_over = True

            # Update bullets
            for bullet in bullets[:]:
                bullet.update()
                if bullet.is_off_screen():
                    bullets.remove(bullet)

            # Check bullet-enemy collisions
            for bullet in bullets[:]:
                hit = False
                for enemy in enemy_birds[:]:
                    if bullet.rect.colliderect(enemy.rect):
                        bullets.remove(bullet)
                        enemy_birds.remove(enemy)
                        score += 5
                        # Add explosion particles
                        for _ in range(8):
                            particles.append(Particle(enemy.x + BIRD_WIDTH // 2, enemy.y + BIRD_HEIGHT // 2,
                                                      (255, 100, 0),
                                                      (random.uniform(-3, 3), random.uniform(-3, 3))))
                        hit = True
                        break

                if not hit:
                    for bee in bees[:]:
                        if bullet.rect.colliderect(bee.rect):
                            bullets.remove(bullet)
                            bees.remove(bee)
                            score += 3
                            # Add bee explosion particles
                            for _ in range(6):
                                particles.append(Particle(bee.x + 12, bee.y + 10,
                                                          (255, 215, 0),
                                                          (random.uniform(-3, 3), random.uniform(-3, 3))))
                            break

        # Update clouds, flowers, stones, and particles
        for cloud in clouds:
            cloud.update()

        for flower in flowers:
            flower.update(pipe_speed)

        for stone in stones:
            stone.update(pipe_speed)

        particles = [p for p in particles if p.life > 0]
        for particle in particles:
            particle.update()

        # Draw everything
        sky_color = sun.get_sky_color()
        for y in range(WINDOW_HEIGHT - 100):
            color_ratio = y / (WINDOW_HEIGHT - 100)
            r = int(sky_color[0] + (255 - sky_color[0]) * color_ratio * 0.3)
            g = int(sky_color[1] + (255 - sky_color[1]) * color_ratio * 0.3)
            b = int(sky_color[2] + (255 - sky_color[2]) * color_ratio * 0.1)
            pygame.draw.line(screen, (r, g, b), (0, y), (WINDOW_WIDTH, y))

        sun.draw(screen)

        for cloud in clouds:
            cloud.draw(screen)

        for pipe in pipes:
            pipe.draw(screen)

        draw_ground(screen, ground_offset)

        for flower in flowers:
            flower.draw(screen)

        for stone in stones:
            stone.draw(screen)

        for particle in particles:
            particle.draw(screen)

        for bullet in bullets:
            bullet.draw(screen)

        for enemy in enemy_birds:
            enemy.draw(screen)

        for bee in bees:
            bee.draw(screen)

        bird.draw(screen)

        # Draw UI
        score_text = big_font.render(str(score), True, WHITE)
        score_outline = big_font.render(str(score), True, BLACK)
        screen.blit(score_outline, (WINDOW_WIDTH // 2 - 12, 52))
        screen.blit(score_text, (WINDOW_WIDTH // 2 - 10, 50))

        difficulty_display = font.render(f"Difficulty: {difficulty.title()}", True, WHITE)
        screen.blit(difficulty_display, (10, 10))
        stage_display = font.render(f"Stage: {stage}", True, WHITE)
        screen.blit(stage_display, (10, 40))

        if game_over:
            overlay = pygame.Surface((WINDOW_WIDTH, WINDOW_HEIGHT))
            overlay.set_alpha(128)
            overlay.fill(BLACK)
            screen.blit(overlay, (0, 0))

            game_over_text = big_font.render("Game Over!", True, WHITE)
            final_score_text = font.render(f"Final Score: {score}", True, WHITE)
            restart_text = font.render("Press DOWN or R to restart", True, WHITE)

            screen.blit(game_over_text, (WINDOW_WIDTH // 2 - 100, WINDOW_HEIGHT // 2 - 80))
            screen.blit(final_score_text, (WINDOW_WIDTH // 2 - 80, WINDOW_HEIGHT // 2 - 30))
            screen.blit(restart_text, (WINDOW_WIDTH // 2 - 120, WINDOW_HEIGHT // 2 + 20))

        pygame.display.flip()
        clock.tick(60)

    pygame.quit()


if __name__ == "__main__":
    main()