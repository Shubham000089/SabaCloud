TB_USERNAME = "//input[@name='j_username']"
TB_PASSWORD = "//input[@name='j_password']"
B_SIGN_IN = "//button[@title='Sign in']"

B_MAIN_MENU = "//button[@title='Main menu']"
L_MY_TEAM = "//div[contains(text(), 'My Team ')]"

D_ORGANISATION_NAME = "//div[@class='sd-team-select-org-container']//span//a"
DO_ORGANISATION = "(//li[@class='sd-team-org-link']) [1]"

EMP_SEARCH_BOX = "//input[@placeholder='Person']"

C_ORGANISATIONS = "//div[@class='sd-team-org-inner-container']//a"

# User Profile
B_BACK = "//span[@class='backlabel']"
L_PLAN = "//a[@title='Plan']"
L_PROFILE = "//a[@title='Profile']"
COMPLETED = "//div[contains(@id, 'container')]//span[@class='sjs-legend-text' and text()='Completed']"
C_COMPLETED = "//div[contains(@id, 'container')]//span[@class='sjs-legend-text' and text()='Completed']/ancestor::li//span[@class='sjs-legend-count']"
# C_COMPLETED = "//table[contains(@id,'sjsgridview')]//tr//b[contains(text(), 'Virtual Classroom') or contains(text(), 'Instructor-Led')] | //table[contains(@id,'sjsgridview')]//tr//a[contains(@class, 'goal')]"
MORE_INFO = "//button[contains(text(), ' More Info ')]"
C_PERSON_NO = "(//span[contains(text(), ' Person No ')]//following::span) [1]"
B_CLOSE = "(//button[contains(@aria-label, 'Close')]) [1]"
PAGE_COUNT_EMP = "//span[contains(@id, 'sjssimplepaginator')]//div//label"

# Certification
FIRST_CERT = "(//a[@title='Course']) [1]"

# Completed
LABEL_TYPE = "//label[contains(text(), 'Type')]"

# Attachment
ATTACHMENT = "//div[contains(@class, 'x-panel')]//span[@class='attachment-link']//a"
B_MODULE_CLOSE = "(//div[@class='x-css-shadow']//following::div[contains(@class, 'x-window')]//div[contains(@class, 'x-window-body')]//span[contains(text(), 'Close')]//..) [1]"

# General
LOADING = "//div[contains(text(), 'Loading..')]"
PLEASE_WAIT = "(//div[contains(text(), 'Please wait')]) [1]"
IFRAME = "//iframe[@title='Main Content']"
LOAD_MORE = "(//div[contains(@Class, 'x-docked-bottom')]//span[contains(text(), 'Load More')]//..) [1]"
B_NEXT = "((//div[contains(@class, 'x-panel')]//span[contains(@class,'x-btn-inner-center') and contains(@id,'paginationNextext')])//..) [1]"
CERTIFICATE_TOOLTIP = "(//div[contains(@class, 'loc-tip-overlay')]//div[@class='loc-tip-content']//div) [1]"