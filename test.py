import folium

#def blackjack_hand_greater_than(hand_1, hand_2):
#     """
#     Return True if hand_1 beats hand_2, and False otherwise.
#
#     In order for hand_1 to beat hand_2 the following must be true:
#     - The total of hand_1 must not exceed 21
#     - The total of hand_1 must exceed the total of hand_2 OR hand_2's total must exceed 21
#
#     Hands are represented as a list of cards. Each card is represented by a string.
#
#     When adding up a hand's total, cards with numbers count for that many points. Face
#     cards ('J', 'Q', and 'K') are worth 10 points. 'A' can count for 1 or 11.
#
#     When determining a hand's total, you should try to count aces in the way that
#     maximizes the hand's total without going over 21. e.g. the total of ['A', 'A', '9'] is 21,
#     the total of ['A', 'A', '9', '3'] is 14.
#
#     Examples:
#     >>> blackjack_hand_greater_than(['K'], ['3', '4'])
#     True
#     >>> blackjack_hand_greater_than(['K'], ['10'])
#     False
#     >>> blackjack_hand_greater_than(['K', 'K', '2'], ['3'])
#     False
#     """
#     dict_card_value = {'1':1, '2':2, '3':3, '4':4, '5':5, '6':6, '7':7, '8':8, '9':9, '10':10,
#                        'J':10, 'Q':10, 'K':10, 'A':11}
#     def calculate_total (hand):
#         hand_total = 0
#         for card in hand:
#             if card in dict_card_value.keys():
#                 hand_total += dict_card_value[card]
#         return hand_total
#
#     hand_1_total = calculate_total(hand_1)
#     hand_2_total = calculate_total(hand_2)
#     if hand_1_total > 21:
#         dict_card_value['A'] = 1
#         hand_1_total = calculate_total(hand_1)
#     if hand_2_total > 21:
#         dict_card_value['A'] = 1
#         hand_2_total = calculate_total(hand_2)
#
#     if hand_1_total > 21:
#         hand_1_total = hand_1_total - 21
#     if hand_2_total > 21:
#         hand_2_total = hand_2_total -21
#     print(hand_1_total, hand_2_total)
#     return hand_1_total > hand_2_total
#
#
# print(blackjack_hand_greater_than(['K'], ['3', '4']))
# print(blackjack_hand_greater_than(['K'], ['10']))
# print(blackjack_hand_greater_than(['K', 'K', '2'], ['3']))
# print(blackjack_hand_greater_than(['9'],['9', 'Q', '8', 'A']))

world_geo = r'C:\Users\mmalikzai\Desktop\san-francisco.geojson'
world_map = folium.Map(location=[37.77,-122.414], zoom_start=2, tiles = 'Mapbox Bright')
