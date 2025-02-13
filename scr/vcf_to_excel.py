# Made for VCF Version: 3.0

import vobject
import pandas as pd

def parse_vcf(vcf_file):
    contacts = []
    error_count = 0

    # Read files with UTF-8 encoding and handle error characters
    with open(vcf_file, encoding="utf-8", errors="ignore") as f:
        vcf_data = f.read()

    # Read files line by line to handle individual parsing errors
    vcards = vcf_data.split("END:VCARD")

    for i, vcard_text in enumerate(vcards, start=1):
        try:
            # Add back "END:VCARD" to keep it in the correct vCard format
            if vcard_text.strip():
                vcard_text += "END:VCARD"
                vcard = next(vobject.readComponents(vcard_text))

                contact = {
                    "Full Name": getattr(vcard, "fn", None).value if hasattr(vcard, "fn") else "",
                    "First Name": "",
                    "Middle Name": "",
                    "Last Name": "",
                    "Prefix": "",
                    "Suffix": "",
                    "Nickname": getattr(vcard, "nickname", None).value if hasattr(vcard, "nickname") else "",
                    "Birthday": getattr(vcard, "bday", None).value if hasattr(vcard, "bday") else "",
                    "Anniversary": getattr(vcard, "anniversary", None).value if hasattr(vcard, "anniversary") else "",
                    "Email Home": "",
                    "Email Work": "",
                    "Telp Cell": "",
                    "Telp Home": "",
                    "Telp Work": "",
                    "Fax Home": "",
                    "Fax Work": "",
                    "Address Home": "",
                    "Address Work": "",
                    "Title": getattr(vcard, "title", None).value if hasattr(vcard, "title") else "",
                    "Organisasi": getattr(vcard, "org", None).value[0] if hasattr(vcard, "org") else "",
                    "URL Work": "",
                    "Note": getattr(vcard, "note", None).value if hasattr(vcard, "note") else "",
                    "URL Facebook": "",
                    "URL Twitter": "",
                    "URL LinkedIn": "",
                    "URL Instagram": "",
                    "URL Youtube": "",
                    "URL Tiktok": ""
                }

                # Splitting names
                if hasattr(vcard, "n"):
                    name_parts = vcard.n.value
                    contact["First Name"] = name_parts.given or ""
                    contact["Middle Name"] = name_parts.additional or ""
                    contact["Last Name"] = name_parts.family or ""
                    contact["Prefix"] = name_parts.prefix or ""
                    contact["Suffix"] = name_parts.suffix or ""

                # Emails
                if hasattr(vcard, "email"):
                    for email in vcard.contents["email"]:
                        email_type = email.params.get("TYPE", [""])[0].lower()
                        if "home" in email_type:
                            contact["Email Home"] = email.value
                        elif "work" in email_type:
                            contact["Email Work"] = email.value

                # Phone Numbers
                if hasattr(vcard, "tel"):
                    for tel in vcard.contents["tel"]:
                        tel_types = [t.lower() for t in tel.params.get("TYPE", [])]  

                        if "cell" in tel_types:
                            contact["Telp Cell"] = tel.value
                        elif "home" in tel_types and "fax" in tel_types:
                            contact["Fax Home"] = tel.value
                        elif "work" in tel_types and "fax" in tel_types:
                            contact["Fax Work"] = tel.value
                        elif "home" in tel_types:
                            contact["Telp Home"] = tel.value
                        elif "work" in tel_types:
                            contact["Telp Work"] = tel.value

                # Addresses
                if hasattr(vcard, "adr"):
                    for adr in vcard.contents["adr"]:
                        adr_type = adr.params.get("TYPE", [""])[0].lower()
                        address = " ".join(filter(None, adr.value)) if isinstance(adr.value, (list, tuple)) else str(adr.value).strip()
                        if "home" in adr_type:
                            contact["Address Home"] = address
                        elif "work" in adr_type:
                            contact["Address Work"] = address

                # URLs
                if hasattr(vcard, "url"):
                    for url in vcard.contents["url"]:
                        url_type = url.params.get("TYPE", [""])[0].lower()
                        if "facebook" in url_type:
                            contact["URL Facebook"] = url.value
                        elif "twitter" in url_type:
                            contact["URL Twitter"] = url.value
                        elif "linkedin" in url_type:
                            contact["URL LinkedIn"] = url.value
                        elif "instagram" in url_type:
                            contact["URL Instagram"] = url.value
                        elif "youtube" in url_type:
                            contact["URL Youtube"] = url.value
                        elif "tiktok" in url_type:
                            contact["URL Tiktok"] = url.value
                        elif "work" in url_type:
                            contact["URL Work"] = url.value

                contacts.append(contact)

        except Exception as e:
            error_count += 1
            print(f"⚠ Peringatan: Gagal memproses kontak ke-{i}. Error: {e}")

    print(f"\n✅ Proses selesai. {len(contacts)} kontak berhasil diproses, {error_count} kontak dilewati.")
    return contacts

# Load VCF file
vcf_file = "contacts.vcf"
contacts = parse_vcf(vcf_file)

# If there are no valid contacts, stop the process
if not contacts:
    print("⛔ Tidak ada kontak yang berhasil diproses. Pastikan file VCF memiliki format yang benar.")
else:
    df = pd.DataFrame(contacts)
    output_file = "contacts.xlsx"
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"\n🎉 Berhasil menyimpan {len(contacts)} kontak ke {output_file}")