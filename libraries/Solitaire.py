import pygame
import random
import sys
import os

# Initialize Pygame
pygame.init()

# Constants
SCREEN_WIDTH = 580  # was 1000
SCREEN_HEIGHT = 450  # was 700
CARD_WIDTH = 60     # was 80
CARD_HEIGHT = 82    # was 110
FOUNDATION_X = 300  # was 400
FOUNDATION_Y = 20
TABLEAU_Y = 150

# Colors
WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
GREEN = (0, 128, 0)
RED = (255, 0, 0)
BLUE = (0, 0, 255)
GRAY = (128, 128, 128)
GOLD = (255, 215, 0)
DARK_GREEN = (34, 139, 34)


class Card:
    def __init__(self, suit, number):
        self.suit = suit  # Diamond, Spades, Hearts, Clubs
        self.number = number  # 1-13
        self.face_up = False
        self.rect = pygame.Rect(0, 0, CARD_WIDTH, CARD_HEIGHT)

    def __str__(self):
        return f"{self.suit} {self.number}"

    def get_display_name(self):
        """Get the display name for the card"""
        if self.number == 1:
            return "A"
        elif self.number == 11:
            return "J"
        elif self.number == 12:
            return "Q"
        elif self.number == 13:
            return "K"
        else:
            return str(self.number)

    def value(self):
        """Return card value for solitaire rules (Ace=1, King=13)"""
        return self.number

    def is_red(self):
        return self.suit in ['Diamond', 'Hearts']

    def can_place_on_foundation(self, foundation_card):
        """Check if this card can be placed on a foundation pile"""
        if foundation_card is None:
            return self.number == 1  # Only Ace can start foundation
        return (self.suit == foundation_card.suit and
                self.number == foundation_card.number + 1)

    def can_place_on_tableau(self, tableau_card):
        """Check if this card can be placed on a tableau pile"""
        if tableau_card is None:
            return self.number == 13  # Only King can go on empty tableau

        # Check colors are different and value is one less
        color_check = self.is_red() != tableau_card.is_red()
        value_check = self.number == tableau_card.number - 1

        return color_check and value_check


class Deck:
    def __init__(self):
        self.cards = []
        suits = ['Diamond', 'Spades', 'Hearts', 'Clubs']

        for suit in suits:
            for number in range(1, 14):  # 1 to 13
                self.cards.append(Card(suit, number))

        self.shuffle()

    def shuffle(self):
        random.shuffle(self.cards)


class SolitaireGame:
    def __init__(self):
        # Reinitialize pygame properly
        if pygame.get_init():
            pygame.quit()
        pygame.init()

        self.screen = pygame.display.set_mode((SCREEN_WIDTH, SCREEN_HEIGHT))
        pygame.display.set_caption("Kherve Solitaire")
        self.clock = pygame.time.Clock()

        # Initialize fonts after pygame.init()
        try:
            self.font_medium = pygame.font.Font(None, 24)
            self.font_large = pygame.font.Font(None, 32)
            self.font_small = pygame.font.Font(None, 16)
        except pygame.error:
            # Fallback if fonts fail to initialize
            pygame.font.init()
            self.font_medium = pygame.font.Font(None, 24)
            self.font_large = pygame.font.Font(None, 32)
            self.font_small = pygame.font.Font(None, 16)

        # Load card images
        self.card_images = {}
        self.card_back_image = None
        self.load_card_images()

        # Game components
        self.deck = Deck()
        self.stock = []
        self.waste = []
        self.foundations = [[] for _ in range(4)]
        self.tableau = [[] for _ in range(7)]

        # Dragging
        self.dragging = False
        self.drag_cards = []
        self.drag_source = None
        self.drag_offset = (0, 0)

        self.setup_game()

    def load_card_images(self):
        """Load card images using your naming convention"""
        suits = ['Diamond', 'Spades', 'Hearts', 'Clubs']

        print("Looking for card images...")

        # Get the directory where this script is located
        current_dir = os.path.dirname(os.path.abspath(__file__))
        cards_dir = os.path.join(current_dir, "cards")

        for suit in suits:
            for number in range(1, 14):
                try:
                    filename = os.path.join(cards_dir, f"{suit} {number}.png")
                    if os.path.exists(filename):
                        image = pygame.image.load(filename)
                        image = pygame.transform.scale(image, (CARD_WIDTH, CARD_HEIGHT))
                        self.card_images[f"{suit}_{number}"] = image
                        print(f"Loaded: {filename}")
                except Exception as e:
                    print(f"Error loading {filename}: {e}")

        # Try to load card back
        try:
            back_path = os.path.join(cards_dir, "back.png")
            if os.path.exists(back_path):
                self.card_back_image = pygame.image.load(back_path)
                self.card_back_image = pygame.transform.scale(self.card_back_image, (CARD_WIDTH, CARD_HEIGHT))
                print("Loaded card back image")
        except Exception as e:
            print(f"Could not load card back: {e}")

        print(f"Total cards loaded: {len(self.card_images)}/52")

    def draw_card(self, screen, card, x, y, face_up=True):
        if x is not None and y is not None:
            rect = pygame.Rect(x, y, CARD_WIDTH, CARD_HEIGHT)
        else:
            rect = card.rect

        if not face_up or not card.face_up:
            # Draw card back
            if self.card_back_image:
                screen.blit(self.card_back_image, rect)
            else:
                # Default card back
                pygame.draw.rect(screen, BLUE, rect)
                pygame.draw.rect(screen, WHITE, rect, 2)
                # Add pattern
                for i in range(5, CARD_WIDTH - 5, 10):
                    for j in range(5, CARD_HEIGHT - 5, 10):
                        pygame.draw.circle(screen, WHITE, (rect.x + i, rect.y + j), 2)
        else:
            # Try to draw card image first
            image_key = f"{card.suit}_{card.number}"
            if image_key in self.card_images:
                screen.blit(self.card_images[image_key], rect)
            else:
                # Fallback to drawing card manually
                pygame.draw.rect(screen, WHITE, rect)
                pygame.draw.rect(screen, BLACK, rect, 2)

                # Card color
                color = RED if card.is_red() else BLACK

                # Display name in top-left
                display_name = card.get_display_name()
                name_text = self.font_medium.render(display_name, True, color)
                screen.blit(name_text, (rect.x + 5, rect.y + 5))

                # Suit symbol in center
                suit_symbols = {
                    'Spades': '♠', 'Hearts': '♥',
                    'Diamond': '♦', 'Clubs': '♣'
                }
                suit_text = self.font_large.render(suit_symbols.get(card.suit, card.suit[0]), True, color)
                suit_rect = suit_text.get_rect(center=(rect.x + CARD_WIDTH // 2, rect.y + CARD_HEIGHT // 2))
                screen.blit(suit_text, suit_rect)

                # Small number in bottom-right
                small_name = self.font_small.render(display_name, True, color)
                screen.blit(small_name, (rect.x + CARD_WIDTH - 20, rect.y + CARD_HEIGHT - 20))

    def draw_empty_space(self, screen, x, y, label=""):
        rect = pygame.Rect(x, y, CARD_WIDTH, CARD_HEIGHT)
        pygame.draw.rect(screen, GRAY, rect)
        pygame.draw.rect(screen, BLACK, rect, 2)

        if label:
            if label in ['Diamond', 'Spades', 'Hearts', 'Clubs']:
                # Draw suit symbol for foundation
                suit_symbols = {
                    'Spades': '♠', 'Hearts': '♥',
                    'Diamond': '♦', 'Clubs': '♣'
                }
                color = RED if label in ['Diamond', 'Hearts'] else BLACK
                text = self.font_large.render(suit_symbols.get(label, label[0]), True, color)
            else:
                text = self.font_small.render(label, True, BLACK)

            text_rect = text.get_rect(center=rect.center)
            screen.blit(text, text_rect)

    def setup_game(self):
        """Deal initial cards for solitaire"""
        # Deal tableau (7 columns, increasing cards per column)
        for col in range(7):
            for row in range(col + 1):
                card = self.deck.cards.pop()
                if row == col:  # Top card face up
                    card.face_up = True
                self.tableau[col].append(card)

        # Remaining cards go to stock
        self.stock = self.deck.cards[:]
        self.waste = []

        self.update_card_positions()

    def update_card_positions(self):
        """Update the screen positions of all cards"""
        # Stock position
        for card in self.stock:
            card.rect.x = 20
            card.rect.y = 20

        # Waste position
        for i, card in enumerate(self.waste):
            card.rect.x = 90 + min(i * 2, 30)  # was 120 + min(i * 2, 40)
            card.rect.y = 20

        # Foundation positions
        for i, foundation in enumerate(self.foundations):
            for card in foundation:
                card.rect.x = FOUNDATION_X + i * (CARD_WIDTH + 8)  # was + 10
                card.rect.y = FOUNDATION_Y

        # Tableau positions
        for col, pile in enumerate(self.tableau):
            for row, card in enumerate(pile):
                card.rect.x = 20 + col * (CARD_WIDTH + 12)  # was + 15
                card.rect.y = TABLEAU_Y + row * 20  # was row * 25

    def get_card_at_pos(self, pos):
        """Find which card was clicked"""
        # Check waste pile (top card only)
        if self.waste and self.waste[-1].rect.collidepoint(pos):
            return self.waste[-1], "waste", len(self.waste) - 1

        # Check foundations
        for i, foundation in enumerate(self.foundations):
            if foundation and foundation[-1].rect.collidepoint(pos):
                return foundation[-1], "foundation", i

        # Check tableau (from bottom to top for proper layering)
        for col, pile in enumerate(self.tableau):
            for row in range(len(pile) - 1, -1, -1):
                if pile[row].rect.collidepoint(pos) and pile[row].face_up:
                    return pile[row], "tableau", (col, row)

        # Check stock area (whether or not there are cards in stock)
        stock_rect = pygame.Rect(20, 20, CARD_WIDTH, CARD_HEIGHT)
        if stock_rect.collidepoint(pos):
            return None, "stock", 0

        return None, None, None

    def can_move_cards(self, cards, target_location, target_index):
        """Check if cards can be moved to target location"""
        if not cards:
            return False

        target_type, target_col = target_location

        if target_type == "foundation":
            # Only single cards to foundation
            if len(cards) != 1:
                return False
            foundation = self.foundations[target_col]
            top_card = foundation[-1] if foundation else None
            return cards[0].can_place_on_foundation(top_card)

        elif target_type == "tableau":
            pile = self.tableau[target_col]
            top_card = pile[-1] if pile else None
            return cards[0].can_place_on_tableau(top_card)

        return False

    def move_cards(self, cards, source_location, target_location, target_index):
        """Move cards from source to target"""
        source_type, source_info = source_location
        target_type, target_col = target_location

        # Remove cards from source
        if source_type == "waste":
            self.waste.pop()
        elif source_type == "foundation":
            self.foundations[source_info].pop()
        elif source_type == "tableau":
            # source_info is a tuple (col, row) for tableau
            if isinstance(source_info, tuple):
                source_col, source_row = source_info
            else:
                source_col = source_info
            source_pile = self.tableau[source_col]
            for _ in range(len(cards)):
                source_pile.pop()
            # Flip next card if needed
            if source_pile and not source_pile[-1].face_up:
                source_pile[-1].face_up = True

        # Add cards to target
        if target_type == "foundation":
            self.foundations[target_col].extend(cards)
        elif target_type == "tableau":
            self.tableau[target_col].extend(cards)

        self.update_card_positions()

    def handle_stock_click(self):
        """Handle clicking on the stock pile"""
        if self.stock:
            # Move card from stock to waste
            card = self.stock.pop()
            card.face_up = True
            self.waste.append(card)
            print(f"Drew {card} from stock")
        else:
            # Reset stock from waste - move all waste cards back to stock
            if self.waste:
                print("Resetting stock from waste pile")
                self.stock = self.waste[:]
                for card in self.stock:
                    card.face_up = False
                self.waste = []
            else:
                print("Both stock and waste are empty - nothing to reset")

        self.update_card_positions()

    def auto_move_to_foundation(self):
        """Try to automatically move cards to foundations"""
        moved = False

        # Check waste pile
        if self.waste:
            card = self.waste[-1]
            for i, foundation in enumerate(self.foundations):
                top_card = foundation[-1] if foundation else None
                if card.can_place_on_foundation(top_card):
                    self.waste.pop()
                    self.foundations[i].append(card)
                    moved = True
                    break

        # Check tableau piles
        if not moved:
            for pile in self.tableau:
                if pile and pile[-1].face_up:
                    card = pile[-1]
                    for i, foundation in enumerate(self.foundations):
                        top_card = foundation[-1] if foundation else None
                        if card.can_place_on_foundation(top_card):
                            pile.pop()
                            self.foundations[i].append(card)
                            if pile and not pile[-1].face_up:
                                pile[-1].face_up = True
                            moved = True
                            break
                    if moved:
                        break

        if moved:
            self.update_card_positions()

        return moved

    def check_win(self):
        """Check if the game is won"""
        return all(len(foundation) == 13 for foundation in self.foundations)

    def run(self):
        running = True
        win_message = ""

        while running:
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    running = False
                elif event.type == pygame.MOUSEBUTTONDOWN:
                    if event.button == 1:  # Left click
                        pos = pygame.mouse.get_pos()
                        card, location, index = self.get_card_at_pos(pos)

                        if location == "stock":
                            self.handle_stock_click()
                        elif card and location in ["waste", "foundation", "tableau"]:
                            # Start dragging
                            self.dragging = True
                            self.drag_source = (location, index)

                            if location == "tableau":
                                col, row = index
                                # Get all cards from this position to end
                                self.drag_cards = self.tableau[col][row:]
                                # Calculate offset relative to the clicked card's actual position
                                clicked_card_y = TABLEAU_Y + row * 20
                                self.drag_offset = (pos[0] - card.rect.x, pos[1] - clicked_card_y)
                            else:
                                self.drag_cards = [card]
                                self.drag_offset = (pos[0] - card.rect.x, pos[1] - card.rect.y)
                    elif event.button == 3:  # Right click - auto move
                        self.auto_move_to_foundation()

                elif event.type == pygame.MOUSEBUTTONUP:
                    if event.button == 1 and self.dragging:
                        pos = pygame.mouse.get_pos()
                        target_found = False

                        # Check foundations
                        for i in range(4):
                            foundation_rect = pygame.Rect(FOUNDATION_X + i * (CARD_WIDTH + 10),
                                                          FOUNDATION_Y, CARD_WIDTH, CARD_HEIGHT)
                            if foundation_rect.collidepoint(pos):
                                if self.can_move_cards(self.drag_cards, ("foundation", i), 0):
                                    source_type, source_index = self.drag_source
                                    self.move_cards(self.drag_cards, (source_type, source_index),
                                                    ("foundation", i), 0)
                                    target_found = True
                                break

                        # Check tableau
                        if not target_found:
                            for col in range(7):
                                tableau_rect = pygame.Rect(20 + col * (CARD_WIDTH + 15), TABLEAU_Y,
                                                           CARD_WIDTH, CARD_HEIGHT + 300)
                                if tableau_rect.collidepoint(pos):
                                    if self.can_move_cards(self.drag_cards, ("tableau", col), 0):
                                        source_type, source_index = self.drag_source
                                        self.move_cards(self.drag_cards, (source_type, source_index),
                                                        ("tableau", col), 0)
                                        target_found = True
                                    break

                        # Reset dragging
                        self.dragging = False
                        self.drag_cards = []
                        self.drag_source = None
                        self.update_card_positions()

                elif event.type == pygame.KEYDOWN:
                    if event.key == pygame.K_n:  # New game
                        self.deck = Deck()
                        self.stock = []
                        self.waste = []
                        self.foundations = [[] for _ in range(4)]
                        self.tableau = [[] for _ in range(7)]
                        self.setup_game()
                        win_message = ""

            # Check for win
            if self.check_win() and not win_message:
                win_message = "Congratulations! You won! Press 'N' for new game."

            # Draw everything
            self.screen.fill(DARK_GREEN)

            # Draw stock - show empty space with "RESET" text when empty but waste has cards
            if self.stock:
                self.draw_card(self.screen, self.stock[-1], 20, 20, False)
            else:
                if self.waste:
                    self.draw_empty_space(self.screen, 20, 20, "RESET")
                else:
                    self.draw_empty_space(self.screen, 20, 20, "STOCK")

            # Draw waste
            if self.waste:
                # Draw up to 3 cards with slight offset
                start_idx = max(0, len(self.waste) - 3)
                for i, card in enumerate(self.waste[start_idx:]):
                    self.draw_card(self.screen, card, 120 + i * 20, 20)
            else:
                self.draw_empty_space(self.screen, 120, 20, "WASTE")

            # Draw foundations
            foundation_suits = ['Diamond', 'Spades', 'Hearts', 'Clubs']
            for i, foundation in enumerate(self.foundations):
                x = FOUNDATION_X + i * (CARD_WIDTH + 10)
                if foundation:
                    self.draw_card(self.screen, foundation[-1], x, FOUNDATION_Y)
                else:
                    self.draw_empty_space(self.screen, x, FOUNDATION_Y, foundation_suits[i])

            # Draw tableau
            for col, pile in enumerate(self.tableau):
                x = 20 + col * (CARD_WIDTH + 15)
                if pile:
                    for row, card in enumerate(pile):
                        if not self.dragging or card not in self.drag_cards:
                            y = TABLEAU_Y + row * 25
                            self.draw_card(self.screen, card, x, y, card.face_up)
                else:
                    self.draw_empty_space(self.screen, x, TABLEAU_Y, "K")

            # Draw dragging cards
            if self.dragging:
                mouse_pos = pygame.mouse.get_pos()
                for i, card in enumerate(self.drag_cards):
                    x = mouse_pos[0] - self.drag_offset[0]
                    y = mouse_pos[1] - self.drag_offset[1] + i * 25
                    self.draw_card(self.screen, card, x, y)

            # Draw UI text
            instructions = self.font_small.render("Left click: Move cards | Right click: Auto-move | N: New game", True,
                                                  WHITE)
            self.screen.blit(instructions, (20, SCREEN_HEIGHT - 40))

            if win_message:
                win_text = self.font_medium.render(win_message, True, GOLD)
                text_rect = win_text.get_rect(center=(SCREEN_WIDTH // 2, SCREEN_HEIGHT - 20))
                self.screen.blit(win_text, text_rect)

            pygame.display.flip()
            self.clock.tick(60)

        pygame.quit()


if __name__ == "__main__":
    game = SolitaireGame()
    game.run()