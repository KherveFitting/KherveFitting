import pygame
import random
import math
import json
import os
from typing import Dict, List, Tuple, Optional

# Initialize Pygame
pygame.init()
pygame.font.init()  # Add this line

# Constants - 2/3 of original size
SCREEN_WIDTH = 800
SCREEN_HEIGHT = 533
FPS = 60

# Colors
BLACK = (0, 0, 0)
WHITE = (255, 255, 255)
GRAY = (128, 128, 128)
LIGHT_GRAY = (200, 200, 200)
DARK_GRAY = (64, 64, 64)
RED = (255, 0, 0)
GREEN = (0, 255, 0)
BLUE = (0, 0, 255)
YELLOW = (255, 255, 0)
ORANGE = (255, 165, 0)
PURPLE = (128, 0, 128)
CYAN = (0, 255, 255)
PINK = (255, 192, 203)
BROWN = (139, 69, 19)
LIGHT_BLUE = (173, 216, 230)
DARK_BLUE = (0, 0, 139)
LIGHT_GREEN = (144, 238, 144)
DARK_GREEN = (0, 100, 0)
SILVER = (192, 192, 192)
GOLD = (255, 215, 0)
VIOLET = (148, 0, 211)
TURQUOISE = (64, 224, 208)

# Element categories and colors
ELEMENT_COLORS = {
    'alkali_metal': (255, 102, 102),
    'alkaline_earth': (255, 222, 173),
    'transition_metal': (255, 192, 203),
    'post_transition': (204, 204, 204),
    'metalloid': (151, 255, 255),
    'nonmetal': (160, 255, 160),
    'halogen': (255, 255, 153),
    'noble_gas': (200, 162, 200),
    'lanthanide': (255, 191, 255),
    'actinide': (255, 153, 204),
    'unknown': (232, 232, 232)
}


class Message:
    def __init__(self, text, color, lifetime=180, y_pos=150):
        self.text = text
        self.color = color
        self.lifetime = lifetime
        self.max_lifetime = lifetime
        self.y_pos = y_pos

    def update(self):
        self.lifetime -= 1

    def draw(self, screen, font):
        if self.lifetime > 0:
            alpha = int(255 * (self.lifetime / self.max_lifetime))
            # Create surface with alpha for fading effect
            text_surface = font.render(self.text, True, self.color)
            fade_surface = pygame.Surface(text_surface.get_size())
            fade_surface.set_alpha(alpha)
            fade_surface.blit(text_surface, (0, 0))

            # Center the text
            text_rect = fade_surface.get_rect(center=(SCREEN_WIDTH // 2, self.y_pos))
            screen.blit(fade_surface, text_rect)

    def is_alive(self):
        return self.lifetime > 0


class Particle:
    def __init__(self, x, y, color, lifetime=60, velocity=None, particle_type="bubble"):
        self.x = x
        self.y = y
        self.start_x = x
        self.start_y = y
        self.color = color
        self.lifetime = lifetime
        self.max_lifetime = lifetime
        self.size = random.randint(2, 6)
        self.particle_type = particle_type

        if velocity:
            self.vel_x, self.vel_y = velocity
        else:
            if particle_type == "bubble":
                # Bubbles rise
                self.vel_x = random.uniform(-0.5, 0.5)
                self.vel_y = random.uniform(-2, -0.5)
            elif particle_type == "gas":
                # Gas particles float around
                angle = random.uniform(0, 2 * math.pi)
                speed = random.uniform(0.5, 2)
                self.vel_x = math.cos(angle) * speed
                self.vel_y = math.sin(angle) * speed - 1  # Slight upward bias
            else:
                # Regular particles
                angle = random.uniform(0, 2 * math.pi)
                speed = random.uniform(1, 4)
                self.vel_x = math.cos(angle) * speed
                self.vel_y = math.sin(angle) * speed

        if particle_type == "gas":
            self.gravity = -0.02  # Gas floats
        elif particle_type == "bubble":
            self.gravity = -0.05  # Bubbles rise
        else:
            self.gravity = 0.1  # Regular particles fall

    def update(self):
        self.x += self.vel_x
        self.y += self.vel_y
        self.vel_y += self.gravity
        self.vel_x *= 0.98  # Air resistance
        self.lifetime -= 1

        # Keep gas particles in bounds
        if self.particle_type == "gas":
            if self.x < 0 or self.x > 800:
                self.vel_x *= -0.5
            if self.y < 0:
                self.vel_y = abs(self.vel_y) * 0.5

    def draw(self, screen):
        if self.lifetime > 0:
            alpha = int(255 * (self.lifetime / self.max_lifetime))
            size = int(self.size * (self.lifetime / self.max_lifetime))
            if size > 0:
                if self.particle_type == "bubble":
                    # Draw bubble with outline
                    pygame.draw.circle(screen, LIGHT_BLUE, (int(self.x), int(self.y)), size)
                    pygame.draw.circle(screen, WHITE, (int(self.x), int(self.y)), size, 1)
                elif self.particle_type == "gas":
                    # Draw gas particle with transparency effect
                    gas_surface = pygame.Surface((size * 2, size * 2))
                    gas_surface.set_alpha(alpha // 2)
                    pygame.draw.circle(gas_surface, self.color, (size, size), size)
                    screen.blit(gas_surface, (int(self.x - size), int(self.y - size)))
                else:
                    pygame.draw.circle(screen, self.color, (int(self.x), int(self.y)), size)

    def is_alive(self):
        return self.lifetime > 0


class Element:
    def __init__(self, symbol, name, atomic_number, category, mass=1):
        self.symbol = symbol
        self.name = name
        self.atomic_number = atomic_number
        self.category = category
        self.mass = mass
        self.color = ELEMENT_COLORS.get(category, GRAY)


class ChemicalReaction:
    def __init__(self, reactants, products, name, energy_change=0, particle_color=WHITE,
                 min_temp=25, max_temp=1000, phase="solid", product_color=WHITE, description="", facts=""):
        self.reactants = reactants  # Dict of element_symbol: count
        self.products = products  # Dict of compound_name: count
        self.name = name
        self.energy_change = energy_change  # Positive = exothermic, Negative = endothermic
        self.particle_color = particle_color
        self.min_temp = min_temp  # Minimum temperature required
        self.max_temp = max_temp  # Maximum safe temperature
        self.phase = phase  # "solid", "liquid", "gas"
        self.product_color = product_color  # Color of the actual product
        self.description = description
        self.facts = facts  # Interesting fact about the compound
        self.discovered = False


class Furnace:
    def __init__(self, x, y, width, height):
        self.rect = pygame.Rect(x, y, width, height)
        self.is_on = False
        self.fuel_level = 0  # 0-10 levels
        self.max_fuel_level = 10
        self.heat_output = 0  # Current heat being generated
        self.flame_particles = []

        # Fuel control buttons
        self.fuel_plus_button = pygame.Rect(x + width + 5, y, 20, 15)
        self.fuel_minus_button = pygame.Rect(x + width + 5, y + 20, 20, 15)

    def increase_fuel(self):
        if self.fuel_level < self.max_fuel_level:
            self.fuel_level += 1

    def decrease_fuel(self):
        if self.fuel_level > 0:
            self.fuel_level -= 1
            if self.fuel_level == 0:
                self.is_on = False

    def toggle(self):
        if self.fuel_level > 0:
            self.is_on = not self.is_on

    def get_temperature_boost(self):
        return self.fuel_level * 100 if self.is_on else 0

    def update(self):
        if self.is_on and self.fuel_level > 0:
            self.heat_output = self.fuel_level * 10  # Heat based on fuel level
            # Create flame particles based on fuel level
            flame_chance = 0.1 + (self.fuel_level * 0.05)
            if random.random() < flame_chance:
                flame_x = self.rect.centerx + random.randint(-10, 10)
                flame_y = self.rect.centery + random.randint(-5, 5)
                flame_intensity = min(255, 100 + self.fuel_level * 15)
                flame_colors = [
                    (flame_intensity, 0, 0),
                    (255, flame_intensity // 2, 0),
                    (255, flame_intensity, 0)
                ]
                flame_color = random.choice(flame_colors)
                self.flame_particles.append(Particle(flame_x, flame_y, flame_color, 30))
        else:
            self.heat_output = 0

        # Update flame particles
        self.flame_particles = [p for p in self.flame_particles if p.is_alive()]
        for particle in self.flame_particles:
            particle.update()

    def draw(self, screen):
        # Draw furnace body with intensity based on fuel level
        if self.is_on:
            intensity = min(255, 100 + self.fuel_level * 15)
            color = (intensity, intensity // 2, 0)
        else:
            color = DARK_GRAY
        pygame.draw.rect(screen, color, self.rect)
        pygame.draw.rect(screen, BLACK, self.rect, 3)

        # Draw flame particles
        for particle in self.flame_particles:
            particle.draw(screen)

        # Draw fuel level indicator
        for i in range(self.max_fuel_level):
            indicator_y = self.rect.bottom - 5 - (i * 3)
            indicator_color = ORANGE if i < self.fuel_level else GRAY
            pygame.draw.rect(screen, indicator_color, (self.rect.x + 2, indicator_y, 8, 2))

        # Draw fuel control buttons
        pygame.draw.rect(screen, GREEN, self.fuel_plus_button)
        pygame.draw.rect(screen, RED, self.fuel_minus_button)
        pygame.draw.rect(screen, BLACK, self.fuel_plus_button, 1)
        pygame.draw.rect(screen, BLACK, self.fuel_minus_button, 1)

        # Draw + and - symbols
        font = pygame.font.Font(None, 16)
        plus_text = font.render("+", True, WHITE)
        minus_text = font.render("-", True, WHITE)
        plus_rect = plus_text.get_rect(center=self.fuel_plus_button.center)
        minus_rect = minus_text.get_rect(center=self.fuel_minus_button.center)
        screen.blit(plus_text, plus_rect)
        screen.blit(minus_text, minus_rect)

        # Draw on/off indicator
        font = pygame.font.Font(None, 16)
        status = "ON" if self.is_on else "OFF"
        status_color = RED if self.is_on else GRAY
        status_surface = font.render(status, True, status_color)
        screen.blit(status_surface, (self.rect.x, self.rect.y - 20))

        # Draw fuel level text
        fuel_text = f"Fuel: {self.fuel_level}"
        fuel_surface = font.render(fuel_text, True, BLACK)
        screen.blit(fuel_surface, (self.rect.x, self.rect.bottom + 5))

        # Draw temperature boost
        temp_boost = self.get_temperature_boost()
        if temp_boost > 0:
            boost_text = f"+{temp_boost}°C"
            boost_surface = font.render(boost_text, True, RED)
            screen.blit(boost_surface, (self.rect.x, self.rect.bottom + 20))


class ReactorVessel:
    def __init__(self, x, y, width, height):
        self.rect = pygame.Rect(x, y, width, height)
        self.contents = {}  # element_symbol: count
        self.base_temperature = 25  # Celsius (room temperature)
        self.pressure = 1.0  # atm
        self.reaction_timer = 0
        self.is_reacting = False
        self.particles = []
        self.liquid_level = 0  # 0-1 ratio for liquid fill
        self.current_phase = "empty"  # "empty", "solid", "liquid", "gas"
        self.current_product_color = WHITE

    def add_element(self, element_symbol, count=1):
        # Clear any previous liquid/gas when adding new elements (fresh experiment)
        if not self.contents:  # If reactor was empty
            self.liquid_level = 0
            self.current_phase = "empty"
            self.particles = []

        if element_symbol in self.contents:
            self.contents[element_symbol] += count
        else:
            self.contents[element_symbol] = count

    def clear(self):
        self.contents = {}
        self.particles = []
        self.is_reacting = False
        self.reaction_timer = 0
        self.liquid_level = 0
        self.current_phase = "empty"

    def get_current_temperature(self, furnace_boost=0):
        return self.base_temperature + furnace_boost

    def can_react(self, reaction, current_temp):
        # Check if we have the right elements
        for reactant, needed_count in reaction.reactants.items():
            if self.contents.get(reactant, 0) < needed_count:
                return False
        # Check if we have exactly the right elements (no extras)
        if len(self.contents) != len(reaction.reactants):
            return False
        # Check temperature requirements
        if current_temp < reaction.min_temp or current_temp > reaction.max_temp:
            return False
        return True

    def has_right_elements(self, reaction):
        # Check if we have the right elements (ignore temperature)
        for reactant, needed_count in reaction.reactants.items():
            if self.contents.get(reactant, 0) < needed_count:
                return False
        # Check if we have exactly the right elements (no extras)
        if len(self.contents) != len(reaction.reactants):
            return False
        return True

    def start_reaction(self, reaction):
        # Remove reactants
        for reactant, count in reaction.reactants.items():
            self.contents[reactant] -= count
            if self.contents[reactant] <= 0:
                del self.contents[reactant]

        self.is_reacting = True
        self.reaction_timer = 120  # 2 seconds at 60 FPS
        self.current_phase = reaction.phase
        self.current_product_color = reaction.product_color

        # Set appropriate visual state based on phase
        if reaction.phase == "liquid":
            self.liquid_level = 0.7  # Fill with liquid
        elif reaction.phase == "gas":
            self.liquid_level = 0  # No liquid for gas
        else:  # solid
            self.liquid_level = 0

        # Create particles
        center_x = self.rect.centerx
        center_y = self.rect.centery

        particle_count = 40 if reaction.phase == "gas" else 30 if reaction.phase == "liquid" else 20
        particle_type = reaction.phase if reaction.phase in ["bubble", "gas"] else "normal"
        if reaction.phase == "liquid":
            particle_type = "bubble"

        for _ in range(particle_count):
            self.particles.append(Particle(
                center_x + random.randint(-20, 20),
                center_y + random.randint(-20, 20),
                reaction.particle_color,
                120 if reaction.phase == "gas" else 90 if reaction.phase == "liquid" else 60,
                particle_type=particle_type
            ))

        return True

    def update(self):
        if self.is_reacting:
            self.reaction_timer -= 1
            if self.reaction_timer <= 0:
                self.is_reacting = False
                # Keep some visual indication of the product
                if self.current_phase == "liquid":
                    self.liquid_level = max(0, self.liquid_level - 0.005)  # Slow evaporation

        # Update particles
        self.particles = [p for p in self.particles if p.is_alive()]
        for particle in self.particles:
            particle.update()

    def draw(self, screen):
        # Draw vessel
        pygame.draw.rect(screen, DARK_GRAY, self.rect, 3)

        # Draw liquid if present
        if self.liquid_level > 0:
            liquid_height = int(self.rect.height * self.liquid_level)
            liquid_rect = pygame.Rect(self.rect.x + 3, self.rect.bottom - liquid_height - 3,
                                      self.rect.width - 6, liquid_height)
            pygame.draw.rect(screen, self.current_product_color, liquid_rect)

        # Draw gas background tint if gas phase
        elif self.current_phase == "gas" and self.is_reacting:
            gas_surface = pygame.Surface((self.rect.width - 6, self.rect.height - 6))
            gas_surface.set_alpha(80)
            gas_surface.fill(self.current_product_color)
            screen.blit(gas_surface, (self.rect.x + 3, self.rect.y + 3))

        # Draw solid powder/crystals at bottom if solid phase
        elif self.current_phase == "solid" and self.is_reacting:
            powder_height = 20
            powder_rect = pygame.Rect(self.rect.x + 3, self.rect.bottom - powder_height - 3,
                                      self.rect.width - 6, powder_height)
            pygame.draw.rect(screen, self.current_product_color, powder_rect)

        # Fill based on contents if no reaction product
        elif self.contents and not self.is_reacting:
            fill_color = LIGHT_GRAY
            fill_rect = pygame.Rect(self.rect.x + 3, self.rect.y + 3,
                                    self.rect.width - 6, self.rect.height - 6)
            pygame.draw.rect(screen, fill_color, fill_rect)

        # Draw contents text
        font = pygame.font.Font(None, 16)
        y_offset = 5
        for element, count in self.contents.items():
            text = f"{element}: {count}"
            text_surface = font.render(text, True, BLACK)
            screen.blit(text_surface, (self.rect.x + 5, self.rect.y + y_offset))
            y_offset += 18

        # Draw particles
        for particle in self.particles:
            particle.draw(screen)


class PeriodicTableButton:
    def __init__(self, element, x, y, width, height):
        self.element = element
        self.rect = pygame.Rect(x, y, width, height)
        self.unlocked = element.symbol in ['H', 'O', 'C', 'N', 'Na', 'Cl', 'Ca', 'Fe']  # Start with basic elements

    def draw(self, screen, font):
        color = self.element.color if self.unlocked else GRAY
        pygame.draw.rect(screen, color, self.rect)
        pygame.draw.rect(screen, BLACK, self.rect, 2)

        if self.unlocked:
            # Element symbol
            symbol_surface = font.render(self.element.symbol, True, BLACK)
            symbol_rect = symbol_surface.get_rect(center=self.rect.center)
            screen.blit(symbol_surface, symbol_rect)

            # Atomic number (small)
            small_font = pygame.font.Font(None, 12)
            num_surface = small_font.render(str(self.element.atomic_number), True, BLACK)
            screen.blit(num_surface, (self.rect.x + 1, self.rect.y + 1))
        else:
            # Locked indicator
            lock_surface = font.render("?", True, DARK_GRAY)
            lock_rect = lock_surface.get_rect(center=self.rect.center)
            screen.blit(lock_surface, lock_rect)


class ChemistryLabGame:
    def __init__(self):
        pygame.init()
        pygame.font.init()

        self.screen = pygame.display.set_mode((SCREEN_WIDTH, SCREEN_HEIGHT))
        pygame.display.set_caption("Material Lab Simulator")
        self.clock = pygame.time.Clock()

        # Fonts
        self.font = pygame.font.Font(None, 28)
        self.small_font = pygame.font.Font(None, 18)
        self.button_font = pygame.font.Font(None, 16)
        self.tiny_font = pygame.font.Font(None, 14)

        # Game state
        self.score = 0
        self.discovered_reactions = 0
        self.total_reactions = 0
        self.show_help = False
        self.help_page = 0
        self.max_help_pages = 3

        # Message system
        self.messages = []

        # Initialize elements and reactions
        self.setup_elements()
        self.setup_reactions()

        # UI Components
        self.reactor = ReactorVessel(300, 200, 150, 120)
        self.furnace = Furnace(470, 220, 60, 80)
        self.setup_periodic_table()

        # Dragging
        self.dragging_element = None
        self.drag_offset = (0, 0)

        # UI buttons
        self.clear_button = pygame.Rect(300, 330, 60, 30)
        self.react_button = pygame.Rect(370, 330, 60, 30)
        self.help_button = pygame.Rect(10, 10, 50, 25)

        # Reaction display
        self.reaction_message = ""
        self.message_timer = 0
        self.message_color = RED
        self.message_position = (300, 150)
        # NEW: Add fact display system
        self.fact_message = ""
        self.fact_timer = 0
        self.fact_position = (550, 400)  # Position in empty space

    def add_message(self, text, color, lifetime=180, y_pos=150):
        self.messages.append(Message(text, color, lifetime, y_pos))

    def setup_elements(self):
        self.elements = {
            'H': Element('H', 'Hydrogen', 1, 'nonmetal'),
            'He': Element('He', 'Helium', 2, 'noble_gas'),
            'Li': Element('Li', 'Lithium', 3, 'alkali_metal'),
            'Be': Element('Be', 'Beryllium', 4, 'alkaline_earth'),
            'B': Element('B', 'Boron', 5, 'metalloid'),
            'C': Element('C', 'Carbon', 6, 'nonmetal'),
            'N': Element('N', 'Nitrogen', 7, 'nonmetal'),
            'O': Element('O', 'Oxygen', 8, 'nonmetal'),
            'F': Element('F', 'Fluorine', 9, 'halogen'),
            'Ne': Element('Ne', 'Neon', 10, 'noble_gas'),
            'Na': Element('Na', 'Sodium', 11, 'alkali_metal'),
            'Mg': Element('Mg', 'Magnesium', 12, 'alkaline_earth'),
            'Al': Element('Al', 'Aluminum', 13, 'post_transition'),
            'Si': Element('Si', 'Silicon', 14, 'metalloid'),
            'P': Element('P', 'Phosphorus', 15, 'nonmetal'),
            'S': Element('S', 'Sulfur', 16, 'nonmetal'),
            'Cl': Element('Cl', 'Chlorine', 17, 'halogen'),
            'Ar': Element('Ar', 'Argon', 18, 'noble_gas'),
            'K': Element('K', 'Potassium', 19, 'alkali_metal'),
            'Ca': Element('Ca', 'Calcium', 20, 'alkaline_earth'),
            'Fe': Element('Fe', 'Iron', 26, 'transition_metal'),
            'Cu': Element('Cu', 'Copper', 29, 'transition_metal'),
            'Zn': Element('Zn', 'Zinc', 30, 'transition_metal'),
        }

    def setup_reactions(self):
        self.reactions = [
            # Basic reactions (room temperature)
            ChemicalReaction({'H': 2, 'O': 1}, {'H2O': 1}, "Water Formation", 10, CYAN,
                             25, 100, "liquid", LIGHT_BLUE, "2H + O → H2O",
                             "Essential for life. Covers 71% of Earth's surface. Boils at 100°C."),

            ChemicalReaction({'Na': 1, 'Cl': 1}, {'NaCl': 1}, "Salt Formation", 15, WHITE,
                             25, 200, "solid", WHITE, "Na + Cl → NaCl",
                             "Table salt. Used for food preservation. Essential for human health."),

            ChemicalReaction({'H': 2, 'Cl': 2}, {'HCl': 2}, "Hydrochloric Acid", 5, YELLOW,
                             25, 150, "liquid", YELLOW, "2H + 2Cl → 2HCl",
                             "Strong acid found in stomach. pH around 1-2. Used in cleaning and metal processing."),

            ChemicalReaction({'Li': 1, 'F': 1}, {'LiF': 1}, "Lithium Fluoride", 12, WHITE,
                             25, 200, "solid", WHITE, "Li + F → LiF",
                             "Used in molten salt reactors. Has the highest melting point of alkali halides (845°C)."),

            ChemicalReaction({'K': 1, 'Cl': 1}, {'KCl': 1}, "Potassium Chloride", 10, WHITE,
                             25, 300, "solid", WHITE, "K + Cl → KCl",
                             "Salt substitute for low-sodium diets. Essential fertilizer for plant growth."),

            ChemicalReaction({'Ca': 1, 'Cl': 2}, {'CaCl2': 1}, "Calcium Chloride", 8, WHITE,
                             25, 250, "solid", WHITE, "Ca + 2Cl → CaCl2",
                             "Road de-icer and desiccant. Absorbs moisture from air. Used in food preservation."),

            # Low temperature reactions (100-300°C)
            ChemicalReaction({'N': 1, 'H': 3}, {'NH3': 1}, "Ammonia Formation", 8, GREEN,
                             100, 300, "gas", LIGHT_GREEN, "N + 3H → NH3",
                             "Key ingredient in fertilizers. Pungent smell. Third most produced chemical worldwide."),

            ChemicalReaction({'S': 1, 'O': 2}, {'SO2': 1}, "Sulfur Dioxide", 12, GRAY,
                             200, 400, "gas", GRAY, "S + 2O → SO2",
                             "Preservative in wine and dried fruits. Major air pollutant. Causes acid rain."),

            ChemicalReaction({'P': 1, 'O': 5}, {'P2O5': 1}, "Phosphorus Pentoxide", 15, WHITE,
                             100, 300, "solid", WHITE, "P + 5O → P2O5",
                             "Powerful dehydrating agent. Used to make phosphoric acid. Extremely hygroscopic."),

            ChemicalReaction({'Na': 3, 'P': 1, 'O': 4}, {'Na3PO4': 1}, "Sodium Phosphate", 18, WHITE,
                             200, 400, "solid", WHITE, "3Na + P + 4O → Na3PO4",
                             "Cleaning agent and food additive. Used in detergents and as meat tenderizer."),

            ChemicalReaction({'Mg': 1, 'Cl': 2}, {'MgCl2': 1}, "Magnesium Chloride", 12, WHITE,
                             150, 350, "solid", WHITE, "Mg + 2Cl → MgCl2",
                             "De-icing salt less corrosive than sodium chloride. Used in tofu production."),

            ChemicalReaction({'Al': 1, 'Cl': 3}, {'AlCl3': 1}, "Aluminum Chloride", 14, WHITE,
                             100, 300, "solid", WHITE, "Al + 3Cl → AlCl3",
                             "Catalyst in organic chemistry. Used in antiperspirants. Sublimes at 180°C."),

            # Medium temperature reactions (300-600°C)
            ChemicalReaction({'C': 1, 'O': 2}, {'CO2': 1}, "Carbon Dioxide", 12, LIGHT_GRAY,
                             300, 600, "gas", LIGHT_GRAY, "C + 2O → CO2",
                             "Greenhouse gas. Plants use it for photosynthesis. Makes soda fizzy. Dry ice at -78°C."),

            ChemicalReaction({'Cu': 1, 'S': 1}, {'CuS': 1}, "Copper Sulfide", 15, DARK_GRAY,
                             400, 700, "solid", DARK_GRAY, "Cu + S → CuS",
                             "Important copper ore called chalcocite. Black mineral used in semiconductors."),

            ChemicalReaction({'Zn': 1, 'S': 1}, {'ZnS': 1}, "Zinc Sulfide", 18, WHITE,
                             500, 800, "solid", WHITE, "Zn + S → ZnS",
                             "Phosphorescent material in glow-in-the-dark products. Major zinc ore called sphalerite."),

            ChemicalReaction({'Al': 2, 'S': 3}, {'Al2S3': 1}, "Aluminum Sulfide", 20, GRAY,
                             400, 700, "solid", GRAY, "2Al + 3S → Al2S3",
                             "Reacts violently with water to produce hydrogen sulfide gas with rotten egg smell."),

            ChemicalReaction({'Be': 1, 'O': 1}, {'BeO': 1}, "Beryllium Oxide", 22, WHITE,
                             400, 700, "solid", WHITE, "Be + O → BeO",
                             "Excellent thermal conductor but toxic. Used in nuclear reactors and electronics."),

            ChemicalReaction({'B': 2, 'O': 3}, {'B2O3': 1}, "Boron Oxide", 16, WHITE,
                             350, 650, "solid", WHITE, "2B + 3O → B2O3",
                             "Glass former used in borosilicate glass. Makes glass heat-resistant like Pyrex."),

            ChemicalReaction({'Li': 2, 'O': 1}, {'Li2O': 1}, "Lithium Oxide", 20, WHITE,
                             300, 600, "solid", WHITE, "2Li + O → Li2O",
                             "Used in ceramic glazes and as CO2 absorber in spacecraft and submarines."),

            ChemicalReaction({'K': 2, 'O': 1}, {'K2O': 1}, "Potassium Oxide", 18, WHITE,
                             350, 650, "solid", WHITE, "2K + O → K2O",
                             "Potash fertilizer component. Creates strongly alkaline solutions in water."),

            # High temperature reactions (600-900°C)
            ChemicalReaction({'Fe': 1, 'S': 1}, {'FeS': 1}, "Iron Sulfide", 18, DARK_GRAY,
                             600, 900, "solid", DARK_GRAY, "Fe + S → FeS",
                             "Pyrite or 'fool's gold' when crystalline. Creates sparks when struck with steel."),

            ChemicalReaction({'Ca': 1, 'O': 1}, {'CaO': 1}, "Calcium Oxide", 20, WHITE,
                             700, 1000, "solid", WHITE, "Ca + O → CaO",
                             "Quicklime used in cement. Reacts violently with water producing heat and steam."),

            ChemicalReaction({'Mg': 1, 'O': 1}, {'MgO': 1}, "Magnesium Oxide", 25, WHITE,
                             600, 900, "solid", WHITE, "Mg + O → MgO",
                             "Refractory material with melting point of 2852°C. Used in furnace linings."),

            ChemicalReaction({'Fe': 2, 'O': 3}, {'Fe2O3': 1}, "Iron Oxide (Rust)", 22, RED,
                             700, 1000, "solid", RED, "2Fe + 3O → Fe2O3",
                             "Common rust. Red pigment in paints and pottery. Mars gets red color from this."),

            ChemicalReaction({'Cu': 1, 'O': 1}, {'CuO': 1}, "Copper Oxide", 16, BLACK,
                             600, 900, "solid", BLACK, "Cu + O → CuO",
                             "Black copper oxide. Used in batteries, ceramics, and as wood preservative."),

            ChemicalReaction({'Zn': 1, 'O': 1}, {'ZnO': 1}, "Zinc Oxide", 18, WHITE,
                             700, 1000, "solid", WHITE, "Zn + O → ZnO",
                             "White pigment in paints and sunscreen. Piezoelectric material used in electronics."),

            ChemicalReaction({'Ca': 1, 'S': 1}, {'CaS': 1}, "Calcium Sulfide", 20, WHITE,
                             800, 1100, "solid", WHITE, "Ca + S → CaS",
                             "Phosphorescent material used in luminous paints and glow-in-the-dark applications."),

            ChemicalReaction({'Mg': 1, 'S': 1}, {'MgS': 1}, "Magnesium Sulfide", 22, WHITE,
                             750, 1050, "solid", WHITE, "Mg + S → MgS",
                             "Used in infrared windows and lenses. Transparent to infrared radiation."),

            # Very high temperature reactions (900°C+)
            ChemicalReaction({'Al': 2, 'O': 3}, {'Al2O3': 1}, "Aluminum Oxide", 30, WHITE,
                             900, 1200, "solid", WHITE, "2Al + 3O → Al2O3",
                             "Sapphire and ruby are crystalline forms. Extremely hard abrasive material."),

            ChemicalReaction({'Si': 1, 'O': 2}, {'SiO2': 1}, "Silicon Dioxide", 25, WHITE,
                             1000, 1200, "solid", WHITE, "Si + 2O → SiO2",
                             "Quartz, sand, and glass. Second most abundant mineral in Earth's crust."),

            ChemicalReaction({'C': 1, 'S': 2}, {'CS2': 1}, "Carbon Disulfide", 8, YELLOW,
                             800, 1100, "liquid", YELLOW, "C + 2S → CS2",
                             "Toxic solvent used in rayon production. Highly flammable with blue flame."),

            ChemicalReaction({'Ca': 1, 'C': 2}, {'CaC2': 1}, "Calcium Carbide", 25, GRAY,
                             900, 1200, "solid", GRAY, "Ca + 2C → CaC2",
                             "Produces acetylene gas when mixed with water. Used in carbide lamps for mining."),

            ChemicalReaction({'Li': 2, 'S': 1}, {'Li2S': 1}, "Lithium Sulfide", 20, WHITE,
                             900, 1200, "solid", WHITE, "2Li + S → Li2S",
                             "Antidepressant effect in small doses. Used in solid-state battery electrolytes."),

            ChemicalReaction({'Be': 1, 'S': 1}, {'BeS': 1}, "Beryllium Sulfide", 24, WHITE,
                             950, 1200, "solid", WHITE, "Be + S → BeS",
                             "Semiconductor material. Extremely toxic like all beryllium compounds."),

            # Additional gas reactions
            ChemicalReaction({'H': 2, 'S': 1}, {'H2S': 1}, "Hydrogen Sulfide", 6, YELLOW,
                             200, 400, "gas", YELLOW, "2H + S → H2S",
                             "Rotten egg smell. Highly toxic gas. Produced by decaying organic matter."),

            ChemicalReaction({'C': 1, 'H': 4}, {'CH4': 1}, "Methane", 10, LIGHT_GRAY,
                             300, 600, "gas", LIGHT_GRAY, "C + 4H → CH4",
                             "Natural gas. Major greenhouse gas. Produced by cows and swamps."),

            ChemicalReaction({'N': 1, 'O': 2}, {'NO2': 1}, "Nitrogen Dioxide", 8, BROWN,
                             400, 700, "gas", BROWN, "N + 2O → NO2",
                             "Brown smog gas. Major air pollutant from car exhaust. Causes acid rain."),

            ChemicalReaction({'S': 1, 'O': 3}, {'SO3': 1}, "Sulfur Trioxide", 14, WHITE,
                             300, 600, "gas", WHITE, "S + 3O → SO3",
                             "Used to make sulfuric acid. Reacts violently with water forming acid mist."),

            ChemicalReaction({'P': 1, 'H': 3}, {'PH3': 1}, "Phosphine", 6, LIGHT_GRAY,
                             200, 500, "gas", LIGHT_GRAY, "P + 3H → PH3",
                             "Highly toxic gas. Spontaneously combustible. Used as semiconductor dopant."),

            ChemicalReaction({'C': 1, 'O': 1}, {'CO': 1}, "Carbon Monoxide", 8, LIGHT_GRAY,
                             400, 800, "gas", LIGHT_GRAY, "C + O → CO",
                             "Silent killer. Odorless toxic gas from incomplete combustion. Binds to hemoglobin."),

            # Complex compounds
            ChemicalReaction({'Ca': 3, 'P': 2, 'O': 8}, {'Ca3P2O8': 1}, "Calcium Phosphate", 28, WHITE,
                             600, 900, "solid", WHITE, "3Ca + 2P + 8O → Ca3P2O8",
                             "Main component of bones and teeth. Used in fertilizers and food supplements."),

            ChemicalReaction({'Mg': 3, 'P': 2, 'O': 8}, {'Mg3P2O8': 1}, "Magnesium Phosphate", 26, WHITE,
                             650, 950, "solid", WHITE, "3Mg + 2P + 8O → Mg3P2O8",
                             "Flame retardant and food additive. Forms strong ceramic materials when fired."),

            ChemicalReaction({'Al': 1, 'P': 1, 'O': 4}, {'AlPO4': 1}, "Aluminum Phosphate", 22, WHITE,
                             500, 800, "solid", WHITE, "Al + P + 4O → AlPO4",
                             "Used in dental cements and as catalyst support. Forms crystals similar to quartz."),

            ChemicalReaction({'Fe': 1, 'P': 1, 'O': 4}, {'FePO4': 1}, "Iron Phosphate", 20, BROWN,
                             600, 900, "solid", BROWN, "Fe + P + 4O → FePO4",
                             "Used in lithium-ion battery cathodes. Rust-resistant coating for steel."),
        ]
        self.total_reactions = len(self.reactions)

    def setup_periodic_table(self):
        self.periodic_buttons = []

        # Compact layout for smaller screen
        elements_layout = [
            ['H', 'He'],
            ['Li', 'Be', 'B', 'C', 'N', 'O', 'F', 'Ne'],
            ['Na', 'Mg', 'Al', 'Si', 'P', 'S', 'Cl', 'Ar'],
            ['K', 'Ca', 'Fe', 'Cu', 'Zn']
        ]

        start_x = 30
        start_y = 50
        button_size = 28
        spacing = 32

        for row_idx, row in enumerate(elements_layout):
            for col_idx, symbol in enumerate(row):
                if symbol in self.elements:
                    x = start_x + col_idx * spacing
                    y = start_y + row_idx * spacing
                    button = PeriodicTableButton(self.elements[symbol], x, y, button_size, button_size)
                    self.periodic_buttons.append(button)

    def handle_element_click(self, pos):
        for button in self.periodic_buttons:
            if button.rect.collidepoint(pos) and button.unlocked:
                self.dragging_element = button.element
                self.drag_offset = (pos[0] - button.rect.centerx, pos[1] - button.rect.centery)
                return True
        return False

    def handle_element_drop(self, pos):
        if self.dragging_element and self.reactor.rect.collidepoint(pos):
            self.reactor.add_element(self.dragging_element.symbol)
            self.dragging_element = None
            return True
        self.dragging_element = None
        return False

    def try_reactions(self):
        current_temp = self.reactor.get_current_temperature(self.furnace.get_temperature_boost())

        # Check if we have any elements first
        if not self.reactor.contents:
            self.show_reaction_message("Add elements to reactor first!", RED, (400, 150))
            return False

        for reaction in self.reactions:
            # Check if we have the right elements
            has_right_elements = True
            for reactant, needed_count in reaction.reactants.items():
                if self.reactor.contents.get(reactant, 0) < needed_count:
                    has_right_elements = False
                    break

            # Check if we have exactly the right elements (no extras)
            has_exact_elements = len(self.reactor.contents) == len(reaction.reactants)

            if has_right_elements and has_exact_elements:
                # Check temperature
                if current_temp < reaction.min_temp:
                    temp_needed = reaction.min_temp - current_temp
                    fuel_needed = max(1, int(temp_needed / 100))
                    message = f"Need more heat! Increase fuel to level {fuel_needed}+ (Need {reaction.min_temp}°C)"
                    self.show_reaction_message(message, ORANGE, (400, 150))
                    return False
                elif current_temp > reaction.max_temp:
                    message = f"Too hot! Reduce fuel (Max {reaction.max_temp}°C)"
                    self.show_reaction_message(message, RED, (400, 150))
                    return False
                else:
                    # Reaction can proceed
                    if self.reactor.start_reaction(reaction):
                        if not reaction.discovered:
                            reaction.discovered = True
                            self.discovered_reactions += 1
                            self.score += 100
                            message = f"New discovery: {reaction.name}! +100 points"
                            self.show_reaction_message(message, GREEN, (400, 150))

                            # Show detailed fact in separate box
                            if hasattr(reaction, 'facts') and reaction.facts:
                                self.show_fact_message(reaction)  # Pass the entire reaction object, not reaction.facts

                            # Unlock new elements based on products
                            self.unlock_elements_from_reaction(reaction)

                        else:
                            self.score += 10
                            message = f"Created {reaction.name} (+10 points)"
                            self.show_reaction_message(message, BLUE, (400, 150))
                        return True

        # No matching reaction found
        elements_list = ", ".join([f"{count}{elem}" for elem, count in self.reactor.contents.items()])
        self.show_reaction_message(f"No reaction found for: {elements_list}", GRAY, (400, 150))
        return False

    def show_reaction_message(self, message, color, position):
        """Display a reaction feedback message"""
        self.reaction_message = message
        self.message_timer = 180  # 3 seconds at 60 FPS
        self.message_color = color
        self.message_position = position

    def show_fact_message(self, reaction):
        """Display a fact with reaction details in a separate box"""
        # Format the complete fact information
        phase_icon = "💧" if reaction.phase == "liquid" else "💨" if reaction.phase == "gas" else "⚗️"

        fact_info = f"{reaction.description}\nPhase: {phase_icon} {reaction.phase.title()}\nFact: {reaction.facts}"

        self.fact_message = fact_info
        self.fact_timer = 1200  # 5 seconds at 60 FPS

    def unlock_elements_from_reaction(self, reaction):
        # Enhanced unlocking logic
        unlock_map = {
            'H2O': ['F', 'Li'],
            'NaCl': ['K', 'Mg'],
            'CO2': ['Si', 'P'],
            'NH3': ['S', 'Ar'],
            'CaO': ['Fe', 'Cu'],
            'FeS': ['Zn', 'Al'],
            'SO2': ['Be', 'B'],
            'CH4': ['He', 'Ne'],
            'LiF': ['Be'],
            'MgO': ['B'],
            'Al2O3': ['Si'],
            'Na3PO4': ['K'],
            'CaC2': ['Si'],
        }

        for product in reaction.products.keys():
            if product in unlock_map:
                for element_symbol in unlock_map[product]:
                    for button in self.periodic_buttons:
                        if button.element.symbol == element_symbol:
                            button.unlocked = True

    def draw_help_panel(self):
        # Semi-transparent overlay
        overlay = pygame.Surface((SCREEN_WIDTH, SCREEN_HEIGHT))
        overlay.set_alpha(200)
        overlay.fill(BLACK)
        self.screen.blit(overlay, (0, 0))

        # Help panel
        panel_rect = pygame.Rect(20, 15, SCREEN_WIDTH - 40, SCREEN_HEIGHT - 30)
        pygame.draw.rect(self.screen, WHITE, panel_rect)
        pygame.draw.rect(self.screen, BLACK, panel_rect, 3)

        # Title with page indicator
        title_text = f"Material Lab HELP - Page {self.help_page + 1}/{self.max_help_pages}"
        title_surface = self.small_font.render(title_text, True, BLACK)
        self.screen.blit(title_surface, (panel_rect.x + 10, panel_rect.y + 10))

        y_offset = 35

        if self.help_page == 0:
            # Page 1: Instructions
            instructions = [
                "HOW TO PLAY:",
                "• Drag elements from periodic table to reactor",
                "• Use +/- buttons to set furnace fuel level (1-10)",
                "• Click furnace to turn on/off heating",
                "• Each fuel level adds 100°C (Level 5 = +500°C)",
                "• Click React when you have right elements + temperature",
                "",
                "REACTION FEEDBACK:",
                "HOT: 'Need more heat' = Increase furnace fuel level",
                "COLD: 'Too hot' = Decrease furnace fuel level",
                "WRONG: 'Wrong elements' = Try different element combination",
                "SUCCESS: Shows interesting facts about the compound!",
                "",
                "REACTION TYPES:",
                "LIQUID: reactions show colored liquid in reactor",
                "GAS: reactions create floating gas particles",
                "SOLID: reactions form powder at bottom",
                "",
                "NAVIGATION:",
                "• Press 1, 2, 3 to switch between help pages",
                "• Press H again to close help",
            ]

            for instruction in instructions:
                color = BLACK
                if instruction.startswith("HOW TO PLAY:") or instruction.startswith(
                        "REACTION FEEDBACK:") or instruction.startswith("REACTION TYPES:") or instruction.startswith(
                    "NAVIGATION:"):
                    color = BLUE
                elif instruction.startswith("•") or instruction.startswith("LIQUID:") or instruction.startswith(
                        "GAS:") or instruction.startswith("SOLID:") or instruction.startswith(
                    "HOT:") or instruction.startswith(
                    "COLD:") or instruction.startswith("WRONG:") or instruction.startswith("SUCCESS:"):
                    color = DARK_GREEN

                inst_surface = self.tiny_font.render(instruction, True, color)
                self.screen.blit(inst_surface, (panel_rect.x + 10, panel_rect.y + y_offset))
                y_offset += 16

        elif self.help_page == 1:
            # Page 2: Basic to Medium Temperature Reactions
            title = "REACTIONS - BASIC TO MEDIUM TEMPERATURE"
            title_surface = self.small_font.render(title, True, BLUE)
            self.screen.blit(title_surface, (panel_rect.x + 10, panel_rect.y + y_offset))
            y_offset += 25

            temp_groups = [
                ("Room Temperature (25°C)", [r for r in self.reactions[:6]]),
                ("Low Heat (100-400°C)", [r for r in self.reactions[6:12]]),
                ("Medium Heat (400-600°C)", [r for r in self.reactions[12:20]]),
            ]

            for group_name, reactions in temp_groups:
                if reactions and y_offset < panel_rect.height - 60:
                    group_surface = self.tiny_font.render(group_name, True, RED)
                    self.screen.blit(group_surface, (panel_rect.x + 15, panel_rect.y + y_offset))
                    y_offset += 18

                    for reaction in reactions[:8]:
                        if y_offset >= panel_rect.height - 40:
                            break

                        phase_icon = "[L]" if reaction.phase == "liquid" else "[G]" if reaction.phase == "gas" else "[S]"
                        # Replace arrow with -> for better compatibility
                        clean_description = reaction.description.replace("→", "->")
                        reaction_line = f"  {phase_icon} {clean_description}"
                        color = GREEN if reaction.discovered else BLACK

                        reaction_surface = self.tiny_font.render(reaction_line, True, color)
                        self.screen.blit(reaction_surface, (panel_rect.x + 20, panel_rect.y + y_offset))
                        y_offset += 14

                    y_offset += 5

        elif self.help_page == 2:
            # Page 3: High Temperature Reactions
            title = "REACTIONS - HIGH TEMPERATURE & EXTREME"
            title_surface = self.small_font.render(title, True, BLUE)
            self.screen.blit(title_surface, (panel_rect.x + 10, panel_rect.y + y_offset))
            y_offset += 25

            temp_groups = [
                ("High Heat (600-900°C)", [r for r in self.reactions[20:28]]),
                ("Very High Heat (900°C+)", [r for r in self.reactions[28:]]),
            ]

            for group_name, reactions in temp_groups:
                if reactions and y_offset < panel_rect.height - 60:
                    group_surface = self.tiny_font.render(group_name, True, RED)
                    self.screen.blit(group_surface, (panel_rect.x + 15, panel_rect.y + y_offset))
                    y_offset += 18

                    for reaction in reactions:
                        if y_offset >= panel_rect.height - 40:
                            break

                        phase_icon = "[L]" if reaction.phase == "liquid" else "[G]" if reaction.phase == "gas" else "[S]"
                        # Replace arrow with -> for better compatibility
                        clean_description = reaction.description.replace("→", "->")
                        reaction_line = f"  {phase_icon} {clean_description}"
                        color = GREEN if reaction.discovered else BLACK

                        reaction_surface = self.tiny_font.render(reaction_line, True, color)
                        self.screen.blit(reaction_surface, (panel_rect.x + 20, panel_rect.y + y_offset))
                        y_offset += 14

                    y_offset += 5

        # Navigation instructions at bottom
        nav_text = "Press 1, 2, 3 for pages | Press H to close"
        nav_surface = self.small_font.render(nav_text, True, RED)
        self.screen.blit(nav_surface, (panel_rect.x + 10, panel_rect.bottom - 25))

    def draw_ui(self):
        # Background
        self.screen.fill(WHITE)

        # Title
        title_surface = self.font.render("Material Lab", True, BLACK)
        self.screen.blit(title_surface, (SCREEN_WIDTH // 2 - title_surface.get_width() // 2, 10))

        # Score and progress
        score_text = f"Score: {self.score}"
        progress_text = f"Discovered: {self.discovered_reactions}/{self.total_reactions}"

        score_surface = self.small_font.render(score_text, True, BLACK)
        progress_surface = self.small_font.render(progress_text, True, BLACK)

        self.screen.blit(score_surface, (SCREEN_WIDTH - 150, 40))
        self.screen.blit(progress_surface, (SCREEN_WIDTH - 150, 60))

        # Help button
        pygame.draw.rect(self.screen, YELLOW, self.help_button)
        pygame.draw.rect(self.screen, BLACK, self.help_button, 2)
        help_text = self.tiny_font.render("HELP", True, BLACK)
        help_rect = help_text.get_rect(center=self.help_button.center)
        self.screen.blit(help_text, help_rect)

        # Periodic table
        for button in self.periodic_buttons:
            button.draw(self.screen, self.button_font)

        # Current temperature display
        current_temp = self.reactor.get_current_temperature(self.furnace.get_temperature_boost())
        temp_color = RED if current_temp > 100 else BLUE if current_temp < 0 else BLACK
        temp_text = f"Current Temp: {current_temp:.0f}°C"
        temp_surface = self.small_font.render(temp_text, True, temp_color)
        self.screen.blit(temp_surface, (300, 175))

        # Reactor vessel
        self.reactor.draw(self.screen)

        # Furnace
        self.furnace.draw(self.screen)

        # Draw reaction feedback message
        if self.reaction_message and self.message_timer > 0:
            # Create semi-transparent background for message
            message_surface = self.small_font.render(self.reaction_message, True, self.message_color)
            message_rect = message_surface.get_rect()
            message_rect.topleft = self.message_position

            # Background rectangle
            bg_rect = message_rect.inflate(20, 10)
            bg_surface = pygame.Surface((bg_rect.width, bg_rect.height))
            bg_surface.set_alpha(200)
            bg_surface.fill(WHITE)
            self.screen.blit(bg_surface, bg_rect)
            pygame.draw.rect(self.screen, BLACK, bg_rect, 2)

            # Message text
            self.screen.blit(message_surface, message_rect)

        # Draw fact message box
        if self.fact_message and self.fact_timer > 0:
            lines = []
            max_width = 250

            for paragraph in self.fact_message.split('\n'):
                if not paragraph.strip():
                    lines.append("")
                    continue

                words = paragraph.split()
                current_line = ""

                for word in words:
                    test_line = (current_line + " " + word) if current_line else word
                    test_surface = self.tiny_font.render(test_line, True, BLACK)

                    if test_surface.get_width() <= max_width:
                        current_line = test_line
                    else:
                        if current_line:
                            lines.append(current_line)
                        current_line = word

                # CRITICAL: Add the final line after processing all words
                if current_line:
                    lines.append(current_line)

            # Calculate box dimensions
            line_height = 15
            box_height = len(lines) * line_height + 40
            box_width = 280

            # Position without bounds checking that might clip content
            fact_x = self.fact_position[0] - 50
            fact_y = self.fact_position[1] - 50

            # Draw box
            fact_rect = pygame.Rect(fact_x, fact_y, box_width, box_height)
            pygame.draw.rect(self.screen, LIGHT_BLUE, fact_rect)
            pygame.draw.rect(self.screen, DARK_BLUE, fact_rect, 2)

            # Header
            header_surface = self.small_font.render("COMPOUND INFO", True, DARK_BLUE)
            self.screen.blit(header_surface, (fact_x + 10, fact_y + 5))

            # Draw all lines
            y_offset = fact_y + 25
            for line in lines:
                if line.strip():
                    color = DARK_GREEN if "→" in line else BLUE if line.startswith(
                        "Phase:") else PURPLE if line.startswith("Fact:") else BLACK
                    line_surface = self.tiny_font.render(line, True, color)
                    self.screen.blit(line_surface, (fact_x + 10, y_offset))
                y_offset += line_height

        # Control buttons
        pygame.draw.rect(self.screen, RED, self.clear_button)
        pygame.draw.rect(self.screen, GREEN, self.react_button)

        clear_text = self.tiny_font.render("Clear", True, WHITE)
        react_text = self.tiny_font.render("React", True, WHITE)

        clear_rect = clear_text.get_rect(center=self.clear_button.center)
        react_rect = react_text.get_rect(center=self.react_button.center)

        self.screen.blit(clear_text, clear_rect)
        self.screen.blit(react_text, react_rect)

        # Instructions
        instructions = [
            "Drag elements to reactor. Use +/- to set furnace fuel level (1-10).",
            "Each level = +100°C. Click furnace to turn on/off. Press H for help."
        ]

        y_offset = 380
        for instruction in instructions:
            inst_surface = self.tiny_font.render(instruction, True, DARK_GRAY)
            self.screen.blit(inst_surface, (30, y_offset))
            y_offset += 15

        # Multi-column discovered reactions list
        reactions_y = 420
        reactions_title = self.small_font.render("Recent Discoveries:", True, BLACK)
        self.screen.blit(reactions_title, (30, reactions_y))

        # Create two columns
        col1_x = 30
        col2_x = 250
        col1_y = reactions_y + 20
        col2_y = reactions_y + 20

        discovered_reactions = [r for r in self.reactions if r.discovered]
        discovered_reactions.reverse()  # Show most recent first

        for i, reaction in enumerate(discovered_reactions[:12]):  # Show up to 12 reactions
            phase_icon = "💧" if reaction.phase == "liquid" else "💨" if reaction.phase == "gas" else "⚗️"
            reaction_text = f"{phase_icon} {reaction.name}"
            reaction_surface = self.tiny_font.render(reaction_text, True, GREEN)

            if i < 6:  # First column
                self.screen.blit(reaction_surface, (col1_x, col1_y))
                col1_y += 14
            else:  # Second column
                self.screen.blit(reaction_surface, (col2_x, col2_y))
                col2_y += 14

        # Draw messages (hints and facts)
        for message in self.messages[:]:
            message.draw(self.screen, self.small_font)

        # Dragging element
        if self.dragging_element:
            mouse_pos = pygame.mouse.get_pos()
            drag_x = mouse_pos[0] - self.drag_offset[0]
            drag_y = mouse_pos[1] - self.drag_offset[1]
            drag_rect = pygame.Rect(drag_x - 15, drag_y - 15, 30, 30)

            pygame.draw.rect(self.screen, self.dragging_element.color, drag_rect)
            pygame.draw.rect(self.screen, BLACK, drag_rect, 2)

            symbol_surface = self.button_font.render(self.dragging_element.symbol, True, BLACK)
            symbol_rect = symbol_surface.get_rect(center=drag_rect.center)
            self.screen.blit(symbol_surface, symbol_rect)

    def run(self):
        running = True
        while running:
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    running = False

                elif event.type == pygame.KEYDOWN:
                    if event.key == pygame.K_h:
                        self.show_help = not self.show_help
                    elif event.key == pygame.K_1:
                        if self.show_help:
                            self.help_page = 0
                    elif event.key == pygame.K_2:
                        if self.show_help:
                            self.help_page = 1
                    elif event.key == pygame.K_3:
                        if self.show_help:
                            self.help_page = 2

                elif event.type == pygame.MOUSEBUTTONDOWN:
                    if event.button == 1:  # Left click
                        pos = pygame.mouse.get_pos()

                        if self.show_help:
                            self.show_help = False
                        elif self.help_button.collidepoint(pos):
                            self.show_help = True
                        elif self.clear_button.collidepoint(pos):
                            self.reactor.clear()
                        elif self.react_button.collidepoint(pos):
                            self.try_reactions()
                        elif self.furnace.fuel_plus_button.collidepoint(pos):
                            self.furnace.increase_fuel()
                        elif self.furnace.fuel_minus_button.collidepoint(pos):
                            self.furnace.decrease_fuel()
                        elif self.furnace.rect.collidepoint(pos):
                            self.furnace.toggle()
                        else:
                            self.handle_element_click(pos)

                elif event.type == pygame.MOUSEBUTTONUP:
                    if event.button == 1 and self.dragging_element:
                        pos = pygame.mouse.get_pos()
                        self.handle_element_drop(pos)

            # Update
            self.reactor.update()
            self.furnace.update()

            # Update message timer
            if self.message_timer > 0:
                self.message_timer -= 1
                if self.message_timer <= 0:
                    self.reaction_message = ""

            # Update fact timer (add this)
            if self.fact_timer > 0:
                self.fact_timer -= 1
                if self.fact_timer <= 0:
                    self.fact_message = ""

            # Update messages
            self.messages = [m for m in self.messages if m.is_alive()]
            for message in self.messages:
                message.update()

            # Draw
            self.draw_ui()

            if self.show_help:
                self.draw_help_panel()

            pygame.display.flip()
            self.clock.tick(FPS)

        pygame.quit()


if __name__ == "__main__":
    game = ChemistryLabGame()
    game.run()