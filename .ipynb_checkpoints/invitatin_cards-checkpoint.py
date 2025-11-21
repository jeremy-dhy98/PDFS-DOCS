def invitation_card(host, date, time, venue, guest):
    return f"""\t\t\t\tDear {guest},
       You are Cordially invited to my
           Graduation ceremony!

            Date: {date}
            Venue: {venue}
            Time: {time}
        
        Looking forward to  Celebrating
        this special day with you.

                Warm regards,
         {host}."""


# Graduation details
graduation_date = "2nd October, 2025"
graduation_venue = "Kitengela Gardens"
graduation_time = "10:00"
host_name = " Mulwa Kitala Jeremiah"

my_people = ["John", "Isaac", "Matthew", "Elizabeth"]

invites = []
for friend in my_people:
    invitation = invitation_card(host_name, graduation_date, graduation_time,
                                graduation_venue, friend)
    invites.append(invitation)
# print("\nNo of invitations: ", len(invites))

