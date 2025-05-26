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
        # Draw enemy bird (red colored)
        bird_surface = pygame.Surface((BIRD_WIDTH, BIRD_HEIGHT), pygame.SRCALPHA)

        # Draw bird body (ellipse) - red enemy
        pygame.draw.ellipse(bird_surface, RED, (0, 0, BIRD_WIDTH, BIRD_HEIGHT))
        pygame.draw.ellipse(bird_surface, (150, 0, 0), (0, 0, BIRD_WIDTH, BIRD_HEIGHT), 3)

        # Draw wing (animated)
        wing_offset = 5 if self.flap_animation > 10 else 0
        pygame.draw.ellipse(bird_surface, (200, 0, 0),
                            (8, 8 - wing_offset, BIRD_WIDTH // 2, BIRD_HEIGHT // 2))

        # Draw eye
        pygame.draw.circle(bird_surface, WHITE, (10, BIRD_HEIGHT // 2 - 3), 6)
        pygame.draw.circle(bird_surface, BLACK, (12, BIRD_HEIGHT // 2 - 3), 3)

        # Draw beak (pointing left)
        beak_points = [(0, BIRD_HEIGHT // 2), (-8, BIRD_HEIGHT // 2 - 2),
                       (-8, BIRD_HEIGHT // 2 + 2)]
        pygame.draw.polygon(bird_surface, (150, 0, 0), beak_points)

        screen.blit(bird_surface, (self.x, self.y))

    def is_off_screen(self):
        return self.x + BIRD_WIDTH < 0


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

        # Create rectangles for collision detection
        self.top_rect = pygame.Rect(x, 0, PIPE_WIDTH, self.gap_y)
        self.bottom_rect = pygame.Rect(x, self.gap_y + pipe_gap, PIPE_WIDTH,
                                       WINDOW_HEIGHT - self.gap_y - pipe_gap - 100)

    def update(self):
        self.x -= self.speed
        self.top_rect.x = self.x
        self.bottom_rect.x = self.x

    def draw(self, screen):
        # Draw pipes
        pygame.draw.rect(screen, PIPE_GREEN, self.top_rect)
        pygame.draw.rect(screen, PIPE_DARK_GREEN, self.top_rect, 4)
        cap_rect = pygame.Rect(self.x - 5, self.gap_y - 20, PIPE_WIDTH + 10, 20)
        pygame.draw.rect(screen, PIPE_GREEN, cap_rect)
        pygame.draw.rect(screen, PIPE_DARK_GREEN, cap_rect, 4)

        pygame.draw.rect(screen, PIPE_GREEN, self.bottom_rect)
        pygame.draw.rect(screen, PIPE_DARK_GREEN, self.bottom_rect, 4)
        cap_rect = pygame.Rect(self.x - 5, self.gap_y + self.gap_size, PIPE_WIDTH + 10, 20)
        pygame.draw.rect(screen, PIPE_GREEN, cap_rect)
        pygame.draw.rect(screen, PIPE_DARK_GREEN, cap_rect, 4)

        pygame.draw.rect(screen, (100, 200, 100),
                         (self.x + 5, 5, 8, self.gap_y - 25))
        pygame.draw.rect(screen, (100, 200, 100),
                         (self.x + 5, self.gap_y + self.gap_size + 25, 8,
                          WINDOW_HEIGHT - self.gap_y - self.gap_size - 125))

        # Draw "Kherve Bird" text randomly on pipes
        if self.show_text:
            font = pygame.font.Font(None, 24)
            text = "Kherve Bird"

            # Randomly choose which pipe section gets text
            pipe_choice = random.choice(['top', 'bottom', 'both'])

            # Top pipe text (rotated 90 degrees)
            if (pipe_choice in ['top', 'both']) and self.gap_y > 120:
                text_surface = font.render(text, True, WHITE)
                rotated_text = pygame.transform.rotate(text_surface, 90)
                text_rect = rotated_text.get_rect()
                text_x = self.x + PIPE_WIDTH // 2 - text_rect.width // 2
                text_y = self.gap_y // 2 - text_rect.height // 2
                screen.blit(rotated_text, (text_x, text_y))

            # Bottom pipe text (rotated 270 degrees)
            if (pipe_choice in ['bottom', 'both']) and WINDOW_HEIGHT - (self.gap_y + self.gap_size) > 220:
                text_surface = font.render(text, True, WHITE)
                rotated_text = pygame.transform.rotate(text_surface, 270)
                text_rect = rotated_text.get_rect()
                text_x = self.x + PIPE_WIDTH // 2 - text_rect.width // 2
                text_y = (self.gap_y + self.gap_size + WINDOW_HEIGHT - 100) // 2 - text_rect.height // 2
                screen.blit(rotated_text, (text_x, text_y))

    def collides_with(self, bird):
        return bird.rect.colliderect(self.top_rect) or bird.rect.colliderect(self.bottom_rect)

    def is_off_screen(self):
        return self.x + PIPE_WIDTH < 0


def draw_ground(screen, ground_offset):
    ground_rect = pygame.Rect(0, WINDOW_HEIGHT - 100, WINDOW_WIDTH, 100)
    pygame.draw.rect(screen, GROUND_BROWN, ground_rect)

    for i in range(0, WINDOW_WIDTH + 20, 20):
        x = (i + ground_offset) % (WINDOW_WIDTH + 20)
        pygame.draw.rect(screen, GRASS_GREEN, (x, WINDOW_HEIGHT - 100, 20, 10))
        for j in range(3):
            grass_x = x + j * 6 + 2
            pygame.draw.line(screen, (0, 100, 0),
                             (grass_x, WINDOW_HEIGHT - 100),
                             (grass_x, WINDOW_HEIGHT - 105), 2)


def show_menu(screen, font, big_font):
    """Show difficulty selection menu"""
    menu_running = True
    selected_difficulty = 'medium'

    while menu_running:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                return None
            elif event.type == pygame.KEYDOWN:
                if event.key == pygame.K_q:  # Add Q to quit from menu too
                    return None
                elif event.key == pygame.K_1:
                    selected_difficulty = 'easy'
                elif event.key == pygame.K_2:
                    selected_difficulty = 'medium'
                elif event.key == pygame.K_3:
                    selected_difficulty = 'hard'
                elif event.key == pygame.K_RETURN:
                    return selected_difficulty

        try:
            # Sky gradient
            for y in range(WINDOW_HEIGHT):
                color_ratio = y / WINDOW_HEIGHT
                r = int(135 + (255 - 135) * color_ratio * 0.3)
                g = int(206 + (255 - 206) * color_ratio * 0.3)
                b = int(235 + (255 - 235) * color_ratio * 0.1)
                pygame.draw.line(screen, (r, g, b), (0, y), (WINDOW_WIDTH, y))

            title_text = big_font.render("Kherve Bird", True, WHITE)
            title_outline = big_font.render("Kherve Bird", True, BLACK)
            screen.blit(title_outline, (WINDOW_WIDTH // 2 - 82, 102))
            screen.blit(title_text, (WINDOW_WIDTH // 2 - 80, 100))

            difficulty_text = font.render("Select Difficulty:", True, WHITE)
            screen.blit(difficulty_text, (WINDOW_WIDTH // 2 - 80, 200))

            easy_color = WHITE if selected_difficulty == 'easy' else (150, 150, 150)
            medium_color = WHITE if selected_difficulty == 'medium' else (150, 150, 150)
            hard_color = WHITE if selected_difficulty == 'hard' else (150, 150, 150)

            easy_text = font.render("1 - Easy", True, easy_color)
            medium_text = font.render("2 - Medium", True, medium_color)
            hard_text = font.render("3 - Hard", True, hard_color)

            screen.blit(easy_text, (WINDOW_WIDTH // 2 - 40, 250))
            screen.blit(medium_text, (WINDOW_WIDTH // 2 - 50, 280))
            screen.blit(hard_text, (WINDOW_WIDTH // 2 - 40, 310))

            start_text = font.render("Press ENTER to start", True, WHITE)
            screen.blit(start_text, (WINDOW_WIDTH // 2 - 90, 380))

            controls_text = font.render("Controls:", True, WHITE)
            screen.blit(controls_text, (50, 450))

            # Create small font for controls
            small_font = pygame.font.Font(None, 24)
            up_text = small_font.render("UP - Jump", True, WHITE)
            space_text = small_font.render("SPACE - Shoot", True, WHITE)
            quit_text = small_font.render("Q - Quit", True, WHITE)

            screen.blit(up_text, (50, 480))
            screen.blit(space_text, (50, 500))
            screen.blit(quit_text, (50, 520))

            pygame.display.flip()

        except pygame.error as e:
            print(f"Rendering error: {e}")
            return None

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
    particles = []
    bullets = []
    enemy_birds = []
    score = 0
    game_over = False
    ground_offset = 0

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
                elif event.key == pygame.K_SPACE and not game_over:
                    bullets.append(bird.shoot())
                elif (event.key == pygame.K_DOWN or event.key == pygame.K_r) and game_over:
                    # Restart with same difficulty
                    bird = Bird(50, WINDOW_HEIGHT // 2)
                    pipes = [Pipe(WINDOW_WIDTH, pipe_speed, pipe_gap)]
                    particles = []
                    bullets = []
                    enemy_birds = []
                    score = 0
                    game_over = False

        if not game_over:
            bird.update()
            ground_offset = (ground_offset - pipe_speed) % 20

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
                    for _ in range(5):
                        particles.append(Particle(bird.x, bird.y,
                                                  (255, 255, 0),
                                                  (random.uniform(-2, 2), random.uniform(-2, 2))))

                if pipe.is_off_screen():
                    pipes.remove(pipe)

            if len(pipes) == 0 or pipes[-1].x < WINDOW_WIDTH - 200:
                pipes.append(Pipe(WINDOW_WIDTH, pipe_speed, pipe_gap))

            # Spawn enemy birds
            if random.random() < enemy_spawn_rate:
                enemy_y = random.randint(50, WINDOW_HEIGHT - 150)
                enemy_birds.append(EnemyBird(WINDOW_WIDTH, enemy_y))

            # Update enemy birds
            for enemy in enemy_birds[:]:
                enemy.update(bird)  # Pass main bird to enemy for opposite movement
                if enemy.is_off_screen():
                    enemy_birds.remove(enemy)
                elif enemy.rect.colliderect(bird.rect):
                    game_over = True

            # Update bullets
            for bullet in bullets[:]:
                bullet.update()
                if bullet.is_off_screen():
                    bullets.remove(bullet)

            # Check bullet-enemy collisions
            for bullet in bullets[:]:
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
                        break

        # Update clouds and particles
        for cloud in clouds:
            cloud.update()

        particles = [p for p in particles if p.life > 0]
        for particle in particles:
            particle.update()

        # Draw everything
        for y in range(WINDOW_HEIGHT - 100):
            color_ratio = y / (WINDOW_HEIGHT - 100)
            r = int(135 + (255 - 135) * color_ratio * 0.3)
            g = int(206 + (255 - 206) * color_ratio * 0.3)
            b = int(235 + (255 - 235) * color_ratio * 0.1)
            pygame.draw.line(screen, (r, g, b), (0, y), (WINDOW_WIDTH, y))

        for cloud in clouds:
            cloud.draw(screen)

        for pipe in pipes:
            pipe.draw(screen)

        draw_ground(screen, ground_offset)

        for particle in particles:
            particle.draw(screen)

        for bullet in bullets:
            bullet.draw(screen)

        for enemy in enemy_birds:
            enemy.draw(screen)

        bird.draw(screen)

        # Draw UI
        score_text = big_font.render(str(score), True, WHITE)
        score_outline = big_font.render(str(score), True, BLACK)
        screen.blit(score_outline, (WINDOW_WIDTH // 2 - 12, 52))
        screen.blit(score_text, (WINDOW_WIDTH // 2 - 10, 50))

        difficulty_display = font.render(f"Difficulty: {difficulty.title()}", True, WHITE)
        screen.blit(difficulty_display, (10, 10))

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