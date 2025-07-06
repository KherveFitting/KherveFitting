import pygame
import random
import sys
import math

# Initialize Pygame
pygame.init()

# Constants
GRID_WIDTH = 10
GRID_HEIGHT = 20
CELL_SIZE = 35
GRID_X_OFFSET = 50
GRID_Y_OFFSET = 50

WINDOW_WIDTH = GRID_WIDTH * CELL_SIZE + 2 * GRID_X_OFFSET + 250
WINDOW_HEIGHT = GRID_HEIGHT * CELL_SIZE + 2 * GRID_Y_OFFSET

# Colors - Enhanced with gradients
BLACK = (0, 0, 0)
WHITE = (255, 255, 255)
GRAY = (128, 128, 128)
DARK_GRAY = (64, 64, 64)
LIGHT_GRAY = (192, 192, 192)

# Enhanced piece colors with gradient information
PIECE_COLORS = [
    {"base": (0, 255, 255), "light": (100, 255, 255), "dark": (0, 200, 200)},  # Cyan - I
    {"base": (0, 0, 255), "light": (100, 100, 255), "dark": (0, 0, 200)},  # Blue - J
    {"base": (255, 165, 0), "light": (255, 200, 100), "dark": (200, 130, 0)},  # Orange - L
    {"base": (255, 255, 0), "light": (255, 255, 100), "dark": (200, 200, 0)},  # Yellow - O
    {"base": (0, 255, 0), "light": (100, 255, 100), "dark": (0, 200, 0)},  # Green - S
    {"base": (128, 0, 128), "light": (180, 100, 180), "dark": (100, 0, 100)},  # Purple - T
    {"base": (255, 0, 0), "light": (255, 100, 100), "dark": (200, 0, 0)}  # Red - Z
]

# Tetris piece shapes
PIECES = [
    # I piece
    [['.....',
      '..#..',
      '..#..',
      '..#..',
      '..#..'],
     ['.....',
      '.....',
      '####.',
      '.....',
      '.....']],

    # J piece
    [['.....',
      '.#...',
      '.###.',
      '.....',
      '.....'],
     ['.....',
      '..##.',
      '..#..',
      '..#..',
      '.....'],
     ['.....',
      '.....',
      '.###.',
      '...#.',
      '.....'],
     ['.....',
      '..#..',
      '..#..',
      '.##..',
      '.....']],

    # L piece
    [['.....',
      '...#.',
      '.###.',
      '.....',
      '.....'],
     ['.....',
      '..#..',
      '..#..',
      '..##.',
      '.....'],
     ['.....',
      '.....',
      '.###.',
      '.#...',
      '.....'],
     ['.....',
      '.##..',
      '..#..',
      '..#..',
      '.....']],

    # O piece
    [['.....',
      '.....',
      '.##..',
      '.##..',
      '.....']],

    # S piece
    [['.....',
      '.....',
      '.##..',
      '##...',
      '.....'],
     ['.....',
      '.#...',
      '.##..',
      '..#..',
      '.....']],

    # T piece
    [['.....',
      '.....',
      '.#...',
      '###..',
      '.....'],
     ['.....',
      '.....',
      '.#...',
      '.##..',
      '.#...'],
     ['.....',
      '.....',
      '.....',
      '###..',
      '.#...'],
     ['.....',
      '.....',
      '.#...',
      '##...',
      '.#...']],

    # Z piece
    [['.....',
      '.....',
      '##...',
      '.##..',
      '.....'],
     ['.....',
      '..#..',
      '.##..',
      '.#...',
      '.....']]
]


def draw_polished_tile(surface, x, y, size, color_info, glow=False):
    """Draw a beautifully polished tile with 3D effect, gradients, and optional glow"""
    base_color = color_info["base"]
    light_color = color_info["light"]
    dark_color = color_info["dark"]

    # Create the main rectangle
    main_rect = pygame.Rect(x, y, size, size)

    # Draw glow effect if enabled
    if glow:
        glow_size = size + 6
        glow_rect = pygame.Rect(x - 3, y - 3, glow_size, glow_size)
        glow_color = tuple(min(255, c + 50) for c in base_color)
        pygame.draw.rect(surface, glow_color, glow_rect, border_radius=4)

        # Inner glow
        inner_glow_size = size + 2
        inner_glow_rect = pygame.Rect(x - 1, y - 1, inner_glow_size, inner_glow_size)
        inner_glow_color = tuple(min(255, c + 30) for c in base_color)
        pygame.draw.rect(surface, inner_glow_color, inner_glow_rect, border_radius=3)

    # Draw the main tile with rounded corners
    pygame.draw.rect(surface, base_color, main_rect, border_radius=3)

    # Create gradient effect by drawing smaller rectangles
    for i in range(5):
        alpha = 0.3 - (i * 0.06)
        if alpha > 0:
            gradient_size = size - (i * 4)
            if gradient_size > 0:
                gradient_rect = pygame.Rect(x + i * 2, y + i * 2, gradient_size, gradient_size)
                gradient_color = tuple(int(base_color[j] + (light_color[j] - base_color[j]) * alpha) for j in range(3))
                pygame.draw.rect(surface, gradient_color, gradient_rect, border_radius=max(1, 3 - i))

    # Add highlight on top-left
    highlight_size = size // 3
    highlight_rect = pygame.Rect(x + 3, y + 3, highlight_size, highlight_size)
    highlight_color = tuple(min(255, c + 80) for c in light_color)
    pygame.draw.rect(surface, highlight_color, highlight_rect, border_radius=2)

    # Add smaller inner highlight
    inner_highlight_size = highlight_size // 2
    inner_highlight_rect = pygame.Rect(x + 5, y + 5, inner_highlight_size, inner_highlight_size)
    inner_highlight_color = tuple(min(255, c + 100) for c in light_color)
    pygame.draw.rect(surface, inner_highlight_color, inner_highlight_rect, border_radius=1)

    # Add shadow on bottom-right
    shadow_thickness = 2
    # Bottom shadow
    bottom_shadow = pygame.Rect(x + shadow_thickness, y + size - shadow_thickness,
                                size - shadow_thickness, shadow_thickness)
    pygame.draw.rect(surface, dark_color, bottom_shadow)

    # Right shadow
    right_shadow = pygame.Rect(x + size - shadow_thickness, y + shadow_thickness,
                               shadow_thickness, size - shadow_thickness)
    pygame.draw.rect(surface, dark_color, right_shadow)

    # Add subtle border
    pygame.draw.rect(surface, tuple(c // 2 for c in dark_color), main_rect, 1, border_radius=3)

    # Add small sparkle effect
    sparkle_x = x + size - 8
    sparkle_y = y + 6
    sparkle_color = tuple(min(255, c + 120) for c in light_color)
    pygame.draw.circle(surface, sparkle_color, (sparkle_x, sparkle_y), 2)
    pygame.draw.circle(surface, WHITE, (sparkle_x, sparkle_y), 1)


def draw_empty_cell(surface, x, y, size):
    """Draw an empty cell with subtle 3D inset effect"""
    cell_rect = pygame.Rect(x, y, size, size)

    # Dark inset background
    pygame.draw.rect(surface, (20, 20, 20), cell_rect)

    # Inner lighter area
    inner_rect = pygame.Rect(x + 1, y + 1, size - 2, size - 2)
    pygame.draw.rect(surface, (35, 35, 35), inner_rect)

    # Subtle grid lines
    pygame.draw.rect(surface, (50, 50, 50), cell_rect, 1)


def draw_background_pattern(surface):
    """Draw a subtle background pattern"""
    for y in range(0, WINDOW_HEIGHT, 20):
        for x in range(0, WINDOW_WIDTH, 20):
            if (x // 20 + y // 20) % 2 == 0:
                pygame.draw.rect(surface, (15, 15, 20), (x, y, 20, 20))
            else:
                pygame.draw.rect(surface, (10, 10, 15), (x, y, 20, 20))


class Piece:
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self.shape = random.randint(0, len(PIECES) - 1)
        self.color = PIECE_COLORS[self.shape]
        self.rotation = 0

    def get_rotated_shape(self):
        return PIECES[self.shape][self.rotation]

    def get_cells(self):
        cells = []
        shape = self.get_rotated_shape()
        for i, row in enumerate(shape):
            for j, cell in enumerate(row):
                if cell == '#':
                    cells.append((self.x + j, self.y + i))
        return cells


class Tetris:
    def __init__(self):
        self.grid = [[None for _ in range(GRID_WIDTH)] for _ in range(GRID_HEIGHT)]
        self.current_piece = self.get_new_piece()
        self.next_piece = self.get_new_piece()
        self.score = 0
        self.lines_cleared = 0
        self.level = 1
        self.fall_time = 0
        self.fall_speed = 500
        self.game_over = False
        self.paused = False  # Add pause state
        self.animation_time = 0

        # Initialize display
        self.screen = pygame.display.set_mode((WINDOW_WIDTH, WINDOW_HEIGHT))
        pygame.display.set_caption("Polished Tetris")
        self.clock = pygame.time.Clock()

        # Load fonts
        self.font = pygame.font.Font(None, 36)
        self.small_font = pygame.font.Font(None, 24)
        self.title_font = pygame.font.Font(None, 48)

    def get_new_piece(self):
        return Piece(GRID_WIDTH // 2 - 2, 0)

    def is_valid_position(self, piece, dx=0, dy=0, rotation=None):
        if rotation is None:
            rotation = piece.rotation

        original_rotation = piece.rotation
        piece.rotation = rotation

        for x, y in piece.get_cells():
            x += dx
            y += dy

            if x < 0 or x >= GRID_WIDTH or y >= GRID_HEIGHT:
                piece.rotation = original_rotation
                return False
            if y >= 0 and self.grid[y][x] is not None:
                piece.rotation = original_rotation
                return False

        piece.rotation = original_rotation
        return True

    def place_piece(self, piece):
        for x, y in piece.get_cells():
            if y >= 0:
                self.grid[y][x] = piece.color

    def clear_lines(self):
        lines_to_clear = []
        for y in range(GRID_HEIGHT):
            if all(cell is not None for cell in self.grid[y]):
                lines_to_clear.append(y)

        for y in lines_to_clear:
            del self.grid[y]
            self.grid.insert(0, [None for _ in range(GRID_WIDTH)])

        lines_cleared = len(lines_to_clear)
        self.lines_cleared += lines_cleared

        if lines_cleared > 0:
            self.score += (lines_cleared ** 2) * 100 * self.level
            self.level = self.lines_cleared // 10 + 1
            self.fall_speed = max(50, 500 - (self.level - 1) * 50)

    def move_piece(self, dx, dy):
        if self.is_valid_position(self.current_piece, dx, dy):
            self.current_piece.x += dx
            self.current_piece.y += dy
            return True
        return False

    def rotate_piece(self):
        num_rotations = len(PIECES[self.current_piece.shape])
        new_rotation = (self.current_piece.rotation + 1) % num_rotations

        if self.is_valid_position(self.current_piece, rotation=new_rotation):
            self.current_piece.rotation = new_rotation

    def drop_piece(self):
        if not self.move_piece(0, 1):
            self.place_piece(self.current_piece)
            self.clear_lines()
            self.current_piece = self.next_piece
            self.next_piece = self.get_new_piece()

            if not self.is_valid_position(self.current_piece):
                self.game_over = True

    def hard_drop(self):
        while self.move_piece(0, 1):
            self.score += 2
        self.drop_piece()

    def draw_grid(self):
        # Draw game area background with border
        board_rect = pygame.Rect(GRID_X_OFFSET - 4, GRID_Y_OFFSET - 4,
                                 GRID_WIDTH * CELL_SIZE + 8, GRID_HEIGHT * CELL_SIZE + 8)

        # Outer glow
        outer_glow = pygame.Rect(GRID_X_OFFSET - 8, GRID_Y_OFFSET - 8,
                                 GRID_WIDTH * CELL_SIZE + 16, GRID_HEIGHT * CELL_SIZE + 16)
        pygame.draw.rect(self.screen, (40, 40, 60), outer_glow, border_radius=8)

        # Main border
        pygame.draw.rect(self.screen, (80, 80, 120), board_rect, border_radius=6)
        pygame.draw.rect(self.screen, (120, 120, 180), board_rect, 2, border_radius=6)

        # Draw the grid cells
        for y in range(GRID_HEIGHT):
            for x in range(GRID_WIDTH):
                cell_x = GRID_X_OFFSET + x * CELL_SIZE
                cell_y = GRID_Y_OFFSET + y * CELL_SIZE

                if self.grid[y][x] is not None:
                    draw_polished_tile(self.screen, cell_x, cell_y, CELL_SIZE, self.grid[y][x])
                else:
                    draw_empty_cell(self.screen, cell_x, cell_y, CELL_SIZE)

    def draw_piece(self, piece, offset_x=0, offset_y=0, glow=False):
        for x, y in piece.get_cells():
            if y >= 0:
                cell_x = GRID_X_OFFSET + (x + offset_x) * CELL_SIZE
                cell_y = GRID_Y_OFFSET + (y + offset_y) * CELL_SIZE
                draw_polished_tile(self.screen, cell_x, cell_y, CELL_SIZE, piece.color, glow)

    def draw_ghost_piece(self):
        ghost_piece = Piece(self.current_piece.x, self.current_piece.y)
        ghost_piece.shape = self.current_piece.shape
        ghost_piece.rotation = self.current_piece.rotation

        while self.is_valid_position(ghost_piece, 0, 1):
            ghost_piece.y += 1

        # Draw ghost with transparency effect
        for x, y in ghost_piece.get_cells():
            if y >= 0:
                cell_x = GRID_X_OFFSET + x * CELL_SIZE
                cell_y = GRID_Y_OFFSET + y * CELL_SIZE

                # Draw subtle ghost outline
                ghost_rect = pygame.Rect(cell_x + 2, cell_y + 2, CELL_SIZE - 4, CELL_SIZE - 4)
                ghost_color = tuple(c // 3 for c in self.current_piece.color["base"])
                pygame.draw.rect(self.screen, ghost_color, ghost_rect, 2, border_radius=2)

    def draw_next_piece(self):
        # Next piece area
        next_x = GRID_X_OFFSET + GRID_WIDTH * CELL_SIZE + 30
        next_y = GRID_Y_OFFSET + 60

        # Background for next piece - make it larger to contain the piece properly
        next_bg = pygame.Rect(next_x - 10, next_y - 40, 120, 120)  # Increased height
        pygame.draw.rect(self.screen, (40, 40, 60), next_bg, border_radius=8)
        pygame.draw.rect(self.screen, (80, 80, 120), next_bg, 2, border_radius=8)

        # Title
        text = self.font.render("Next", True, (200, 200, 255))
        self.screen.blit(text, (next_x, next_y - 35))

        # Calculate piece bounds to center it properly
        shape = self.next_piece.get_rotated_shape()
        min_x, max_x = 5, 0
        min_y, max_y = 5, 0

        for i, row in enumerate(shape):
            for j, cell in enumerate(row):
                if cell == '#':
                    min_x = min(min_x, j)
                    max_x = max(max_x, j)
                    min_y = min(min_y, i)
                    max_y = max(max_y, i)

        # Center the piece in the next box
        piece_width = max_x - min_x + 1
        piece_height = max_y - min_y + 1
        center_offset_x = (5 - piece_width) // 2 - min_x
        center_offset_y = (4 - piece_height) // 2 - min_y

        # Draw next piece with smaller tiles, centered in the box
        for i, row in enumerate(shape):
            for j, cell in enumerate(row):
                if cell == '#':
                    tile_x = next_x + (j + center_offset_x) * 18
                    tile_y = next_y + (i + center_offset_y) * 18
                    draw_polished_tile(self.screen, tile_x, tile_y, 16, self.next_piece.color)

    def draw_ui(self):
        ui_x = GRID_X_OFFSET + GRID_WIDTH * CELL_SIZE + 30
        ui_y = GRID_Y_OFFSET + 200  # Moved down to accommodate larger next piece box

        # UI Background
        ui_bg = pygame.Rect(ui_x - 10, ui_y - 10, 180, 200)
        pygame.draw.rect(self.screen, (40, 40, 60), ui_bg, border_radius=8)
        pygame.draw.rect(self.screen, (80, 80, 120), ui_bg, 2, border_radius=8)

        # Score with glow effect
        score_text = self.font.render(f"Score", True, (200, 200, 255))
        score_value = self.font.render(f"{self.score}", True, (255, 255, 255))
        self.screen.blit(score_text, (ui_x, ui_y))
        self.screen.blit(score_value, (ui_x, ui_y + 25))

        # Lines
        lines_text = self.font.render(f"Lines", True, (200, 200, 255))
        lines_value = self.font.render(f"{self.lines_cleared}", True, (255, 255, 255))
        self.screen.blit(lines_text, (ui_x, ui_y + 60))
        self.screen.blit(lines_value, (ui_x, ui_y + 85))

        # Level
        level_text = self.font.render(f"Level", True, (200, 200, 255))
        level_value = self.font.render(f"{self.level}", True, (255, 255, 255))
        self.screen.blit(level_text, (ui_x, ui_y + 120))
        self.screen.blit(level_value, (ui_x, ui_y + 145))

    def draw_controls(self):
        controls_x = GRID_X_OFFSET + GRID_WIDTH * CELL_SIZE + 30
        controls_y = GRID_Y_OFFSET + 420  # Moved down to accommodate larger UI

        # Controls background
        controls_bg = pygame.Rect(controls_x - 10, controls_y - 10, 180, 180)  # Increased height
        pygame.draw.rect(self.screen, (40, 40, 60), controls_bg, border_radius=8)
        pygame.draw.rect(self.screen, (80, 80, 120), controls_bg, 2, border_radius=8)

        controls = [
            "Controls:",
            "A/D - Move",
            "S - Soft drop",
            "W - Rotate",
            "Space - Hard drop",
            "P - Pause",  # Added pause control
            "Q/Esc - Quit",  # Added quit control
            "R - Restart"
        ]

        for i, control in enumerate(controls):
            color = (200, 200, 255) if i == 0 else (180, 180, 200)
            font = self.font if i == 0 else self.small_font
            text = font.render(control, True, color)
            self.screen.blit(text, (controls_x, controls_y + i * 20))  # Reduced spacing

    def draw(self):
        # Animated background
        draw_background_pattern(self.screen)

        self.draw_grid()

        # Only draw ghost piece and current piece if not paused
        if not self.paused:
            self.draw_ghost_piece()
            # Draw current piece with glow animation
            glow = math.sin(self.animation_time * 0.01) > 0.5
            self.draw_piece(self.current_piece, glow=glow)

        self.draw_next_piece()
        self.draw_ui()
        self.draw_controls()

        # Title
        title_text = self.title_font.render("TETRIS", True, (255, 255, 255))
        title_shadow = self.title_font.render("TETRIS", True, (100, 100, 100))
        self.screen.blit(title_shadow, (GRID_X_OFFSET + 2, 12))
        self.screen.blit(title_text, (GRID_X_OFFSET, 10))

        # Pause overlay
        if self.paused:
            overlay = pygame.Surface((WINDOW_WIDTH, WINDOW_HEIGHT))
            overlay.set_alpha(128)
            overlay.fill((0, 0, 0))
            self.screen.blit(overlay, (0, 0))

            pause_text = self.title_font.render("PAUSED", True, (255, 255, 100))
            continue_text = self.font.render("Press P to continue", True, (200, 200, 200))

            text_rect = pause_text.get_rect(center=(WINDOW_WIDTH // 2, WINDOW_HEIGHT // 2))
            continue_rect = continue_text.get_rect(center=(WINDOW_WIDTH // 2, WINDOW_HEIGHT // 2 + 60))

            # Text shadows
            shadow_offset = 3
            pause_shadow = self.title_font.render("PAUSED", True, (100, 100, 0))
            continue_shadow = self.font.render("Press P to continue", True, (100, 100, 100))

            shadow_rect = text_rect.copy()
            shadow_rect.x += shadow_offset
            shadow_rect.y += shadow_offset
            continue_shadow_rect = continue_rect.copy()
            continue_shadow_rect.x += shadow_offset
            continue_shadow_rect.y += shadow_offset

            self.screen.blit(pause_shadow, shadow_rect)
            self.screen.blit(continue_shadow, continue_shadow_rect)
            self.screen.blit(pause_text, text_rect)
            self.screen.blit(continue_text, continue_rect)

        # Game over overlay
        if self.game_over:
            overlay = pygame.Surface((WINDOW_WIDTH, WINDOW_HEIGHT))
            overlay.set_alpha(128)
            overlay.fill((0, 0, 0))
            self.screen.blit(overlay, (0, 0))

            game_over_text = self.title_font.render("GAME OVER", True, (255, 100, 100))
            restart_text = self.font.render("Press R to restart", True, (200, 200, 200))

            text_rect = game_over_text.get_rect(center=(WINDOW_WIDTH // 2, WINDOW_HEIGHT // 2))
            restart_rect = restart_text.get_rect(center=(WINDOW_WIDTH // 2, WINDOW_HEIGHT // 2 + 60))

            # Text shadows
            shadow_offset = 3
            game_over_shadow = self.title_font.render("GAME OVER", True, (100, 0, 0))
            restart_shadow = self.font.render("Press R to restart", True, (100, 100, 100))

            shadow_rect = text_rect.copy()
            shadow_rect.x += shadow_offset
            shadow_rect.y += shadow_offset
            restart_shadow_rect = restart_rect.copy()
            restart_shadow_rect.x += shadow_offset
            restart_shadow_rect.y += shadow_offset

            self.screen.blit(game_over_shadow, shadow_rect)
            self.screen.blit(restart_shadow, restart_shadow_rect)
            self.screen.blit(game_over_text, text_rect)
            self.screen.blit(restart_text, restart_rect)

        pygame.display.flip()

    def reset_game(self):
        self.grid = [[None for _ in range(GRID_WIDTH)] for _ in range(GRID_HEIGHT)]
        self.current_piece = self.get_new_piece()
        self.next_piece = self.get_new_piece()
        self.score = 0
        self.lines_cleared = 0
        self.level = 1
        self.fall_time = 0
        self.fall_speed = 500
        self.game_over = False
        self.paused = False  # Reset pause state

    def run(self):
        running = True

        while running:
            dt = self.clock.tick(60)

            # Only update animation and fall time if not paused
            if not self.paused:
                self.fall_time += dt
                self.animation_time += dt

            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    running = False
                elif event.type == pygame.KEYDOWN:
                    # Quit keys (work anytime)
                    if event.key == pygame.K_q or event.key == pygame.K_ESCAPE:
                        running = False
                    # Restart key (work anytime)
                    elif event.key == pygame.K_r:
                        self.reset_game()
                    # Pause key (only when not game over)
                    elif event.key == pygame.K_p and not self.game_over:
                        self.paused = not self.paused
                    # Game controls (only when not paused and not game over)
                    elif not self.paused and not self.game_over:
                        if event.key == pygame.K_a or event.key == pygame.K_LEFT:
                            self.move_piece(-1, 0)
                        elif event.key == pygame.K_d or event.key == pygame.K_RIGHT:
                            self.move_piece(1, 0)
                        elif event.key == pygame.K_s or event.key == pygame.K_DOWN:
                            self.move_piece(0, 1)
                        elif event.key == pygame.K_w or event.key == pygame.K_UP:
                            self.rotate_piece()
                        elif event.key == pygame.K_SPACE:
                            self.hard_drop()

            # Game logic (only when not paused and not game over)
            if not self.paused and not self.game_over:
                keys = pygame.key.get_pressed()
                if keys[pygame.K_s]:
                    if self.fall_time >= self.fall_speed // 10:
                        self.drop_piece()
                        self.fall_time = 0
                        self.score += 1
                elif self.fall_time >= self.fall_speed:
                    self.drop_piece()
                    self.fall_time = 0

            self.draw()

        pygame.quit()
        # sys.exit()


if __name__ == "__main__":
    game = Tetris()
    game.run()