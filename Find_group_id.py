import vk_api

vk = vk_api.VkApi(
    token="vk1.a.N6k9S9XkHfGvoAxZf6b8NGBRm-6EKftc3GjUsCBgvLe8rjbdbC_nXb4c2Mx18Z8fUvdjKaDqYkyrZkJvlAPV2fIuerW3FToUbOTHctI_dPAtG_yU6--EnNuNRoMu0dbOrQL9V07FX4ML3FWWXIOT2k-a_lMOwvVw8Eg_oNZ64asQ9ptqTA3BG6vc1CwuFiRLN4CufB5AhtmaiUqirGRKzw")
all_convs = vk.method("messages.getConversations", {"fields": "name"})

# for chat in all_convs['items']:
#     lol = chat['conversation']['peer']['type']
#     if lol != "user":
#         print(chat['conversation']['peer']['id'], chat['conversation']['chat_settings']['title'])

all_peeps = vk.method("messages.getConversationMembers", {"peer_id": 2000000070})
# print(all_peeps['profiles'][0].keys())
for person in all_peeps['profiles']:
    first_name = person["first_name"]
    last_name = person["last_name"]
    print(f'{first_name} {last_name}', end=',')
