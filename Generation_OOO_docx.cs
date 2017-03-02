using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;

namespace OnCore
{
    class Generation_OOO_docx : IGenerate
    {
 
        private string registering_authority;
        private string soglas;
        private string name_firma_ru;
        private string name_firma_by;
        private string oked;
        private string oked_name;
        private string all_fond;
        private string date_protocol;
        private string date_protocol_2;
        private string index_ur;
        private string region_ur;
        private string district_ur;

        private string street_ur;
        private string house_ur;
        private string office_ur;
        private string korpus_ur;
        private string locality_ur;
        private string name_city_ur;
        private  int quantity_founder;

        //Step 3
        private string firstname_director;
        private string name_director;
        private string lastname_director;
        private string datebirthd_director;
        private string placebirthd_director;
        private string index_director;
        private string region_director;
        private string district_director;
        private string locality_director;
        private string name_city_director;
        private string street_director;
        private string korpus_director;
        private string house_director;
        private string app_director;
        private string idpassport_director;
        private string numberpassport_director;
        private string datapassport_director;
        private string validpassport_director;
        private string authority_passport_director;
        private string code_director;
        private string phone_director;
        private string director_country;
        private string director_citizen; // Добавить

        private string director_type_document;
        //Step 4
        private string founder_1_firstname;
        private string founder_1_name;
        private string founder_1_lastname;
        private string founder_1_sex;
        private string founder_1_datebirthd;
        private string founder_1_placebirthd;
        private string founder_1_index;
        private string founder_1_region;
        private string founder_1_district;
        private string locality_founder_1;
        private string founder_1_name_city;
        private string founder_1_street;
        private string founder_1_korpus;
        private string founder_1_number_house;
        private string founder_1_app;
        private string founder_1_id_passport;
        private string founder_1_number_passport;
        private string founder_1_date_passport;
        private string founder_1_valid_passport;
        private string founder_1_registration_passport;
        private string founder_1_code;
        private string founder_1_phone;
        private string founder_1_fond;
        private string founder_1_precent_fond;
        private string founder_1_type_document;
        private string founder_1_citizen;
        private string founder_1_country;
        //Step 5
        private string founder_2_firstname;
        private string founder_2_name;
        private string founder_2_lastname;
        private string founder_2_sex;
        private string founder_2_datebirthd;
        private string founder_2_placebirthd;
        private string founder_2_index;
        private string founder_2_region;
        private string founder_2_district;
        private string locality_founder_2;
        private string founder_2_name_city;
        private string founder_2_street;
        private string founder_2_korpus;
        private string founder_2_number_house;
        private string founder_2_app;
        private string founder_2_id_passport;
        private string founder_2_number_passport;
        private string founder_2_date_passport;
        private string founder_2_valid_passport;
        private string founder_2_registration_passport;
        private string founder_2_code;
        private string founder_2_phone;
        private string founder_2_fond;
        private string founder_2_precent_fond;
        private string founder_2_country;
        private string founder_2_citizen;

        private string founder_2_type_document;
        // Step 6//

        private string founder_3_firstname;
        private string founder_3_name;
        private string founder_3_lastname;
        private string founder_3_sex;
        private string founder_3_datebirthd;
        private string founder_3_placebirthd;
        private string founder_3_index;
        private string founder_3_region;
        private string founder_3_district;
        private string locality_founder_3;
        private string founder_3_name_city;
        private string founder_3_street;
        private string founder_3_korpus;
        private string founder_3_number_house;
        private string founder_3_app;
        private string founder_3_id_passport;
        private string founder_3_number_passport;
        private string founder_3_date_passport;
        private string founder_3_valid_passport;
        private string founder_3_registration_passport;
        private string founder_3_code;
        private string founder_3_phone;
        private string founder_3_fond;
        private string founder_3_precent_fond;
        private string founder_3_country;
        private string founder_3_citizen;

        private string founder_3_type_document;

        public Generation_OOO_docx(IDictionary<string, string> generation_ooo_dictionary)
        {

           
            registering_authority = generation_ooo_dictionary["registering_authority"];
            soglas = generation_ooo_dictionary["soglas"];
            name_firma_ru = generation_ooo_dictionary["name_firma_ru"];
            name_firma_by = generation_ooo_dictionary["name_firma_by"];
            oked = generation_ooo_dictionary["oked"];
            oked_name = generation_ooo_dictionary["oked_name"];
            all_fond = generation_ooo_dictionary["all_fond"];
            date_protocol = generation_ooo_dictionary["date_protocol"];
            date_protocol_2 = generation_ooo_dictionary["date_protocol_2"];
            index_ur = generation_ooo_dictionary["index_ur"];
            region_ur = generation_ooo_dictionary["region_ur"];
            district_ur = generation_ooo_dictionary["district_ur"];
            street_ur = generation_ooo_dictionary["street_ur"];
            house_ur = generation_ooo_dictionary["house_ur"];
            office_ur = generation_ooo_dictionary["office_ur"];
            korpus_ur = generation_ooo_dictionary["korpus_ur"];
            locality_ur = generation_ooo_dictionary["locality_ur"];
            name_city_ur = generation_ooo_dictionary["name_city_ur"];
            quantity_founder = Int32.Parse(generation_ooo_dictionary["quantity_founder"]);




            //Step 3
            firstname_director = generation_ooo_dictionary["firstname_director"];
            name_director = generation_ooo_dictionary["name_director"];
            lastname_director = generation_ooo_dictionary["lastname_director"];
            datebirthd_director = generation_ooo_dictionary["datebirthd_director"];
            placebirthd_director = generation_ooo_dictionary["placebirthd_director"];
            index_director = generation_ooo_dictionary["index_director"];
            region_director = generation_ooo_dictionary["region_director"];
            district_director = generation_ooo_dictionary["district_director"];
            locality_director = generation_ooo_dictionary["locality_director"];
            name_city_director = generation_ooo_dictionary["name_city_director"];
            street_director = generation_ooo_dictionary["street_director"];
            korpus_director = generation_ooo_dictionary["korpus_director"];
            house_director = generation_ooo_dictionary["house_director"];
            app_director = generation_ooo_dictionary["app_director"];
            idpassport_director = generation_ooo_dictionary["idpassport_director"];
            numberpassport_director = generation_ooo_dictionary["numberpassport_director"];
            datapassport_director = generation_ooo_dictionary["datapassport_director"];
            validpassport_director = generation_ooo_dictionary["validpassport_director"];
            authority_passport_director = generation_ooo_dictionary["authority_passport_director"];
            code_director = generation_ooo_dictionary["code_director"];
            phone_director = generation_ooo_dictionary["phone_director"];
            director_country = generation_ooo_dictionary["director_country"];
            director_citizen = generation_ooo_dictionary["director_citizen"];

            director_type_document = generation_ooo_dictionary["director_type_document"];
            switch (director_type_document)
            {
                case "passport_director":
                    director_type_document = "Паспорт";
                    break;
                case "residence_director":
                    director_type_document = "Вид на жительство";
                    break;
            }
            //Step 4

            switch (quantity_founder)
            {
                case 1:
                    founder_1_firstname = generation_ooo_dictionary["founder_1_firstname"];
                    founder_1_name = generation_ooo_dictionary["founder_1_name"];
                    founder_1_lastname = generation_ooo_dictionary["founder_1_lastname"];
                    founder_1_sex = generation_ooo_dictionary["founder_1_sex"];
                    founder_1_datebirthd = generation_ooo_dictionary["founder_1_datebirthd"];
                    founder_1_placebirthd = generation_ooo_dictionary["founder_1_placebirthd"];
                    founder_1_index = generation_ooo_dictionary["founder_1_index"];
                    founder_1_region = generation_ooo_dictionary["founder_1_region"];
                    founder_1_district = generation_ooo_dictionary["founder_1_district"];
                    locality_founder_1 = generation_ooo_dictionary["locality_founder_1"];
                    founder_1_name_city = generation_ooo_dictionary["founder_1_name_city"];
                    founder_1_street = generation_ooo_dictionary["founder_1_street"];
                    founder_1_korpus = generation_ooo_dictionary["founder_1_korpus"];
                    founder_1_number_house = generation_ooo_dictionary["founder_1_number_house"];
                    founder_1_app = generation_ooo_dictionary["founder_1_app"];
                    founder_1_id_passport = generation_ooo_dictionary["founder_1_id_passport"];
                    founder_1_number_passport = generation_ooo_dictionary["founder_1_number_passport"];
                    founder_1_date_passport = generation_ooo_dictionary["founder_1_date_passport"];
                    founder_1_valid_passport = generation_ooo_dictionary["founder_1_valid_passport"];
                    founder_1_registration_passport = generation_ooo_dictionary["founder_1_registration_passport"];
                    founder_1_code = generation_ooo_dictionary["founder_1_code"];
                    founder_1_phone = generation_ooo_dictionary["founder_1_phone"];
                    founder_1_fond = generation_ooo_dictionary["founder_1_fond"];
                    founder_1_precent_fond = generation_ooo_dictionary["founder_1_precent_fond"];
                    founder_1_type_document = generation_ooo_dictionary["founder_1_type_document"];
                    founder_1_citizen = generation_ooo_dictionary["founder_1_citizen"];
                    founder_1_country = generation_ooo_dictionary["founder_1_country"];
                    switch (founder_1_type_document)
                    {
                        case "founder_1_passport":
                            founder_1_type_document = "Паспорт";
                            break;
                        case "founder_1_residence":
                            founder_1_type_document = "Вид на жительство";
                            break;
                    }
                    switch (founder_1_sex)
                    {
                        case "founder_1_sex_male":
                            founder_1_sex = "Мужской";
                            break;
                        case "founder_1_sex_female":
                            founder_1_sex = "Женский";
                            break;
                    }
                    break;
                case 2:
                    founder_1_firstname = generation_ooo_dictionary["founder_1_firstname"];
                    founder_1_name = generation_ooo_dictionary["founder_1_name"];
                    founder_1_lastname = generation_ooo_dictionary["founder_1_lastname"];
                    founder_1_sex = generation_ooo_dictionary["founder_1_sex"];
                    founder_1_datebirthd = generation_ooo_dictionary["founder_1_datebirthd"];
                    founder_1_placebirthd = generation_ooo_dictionary["founder_1_placebirthd"];
                    founder_1_index = generation_ooo_dictionary["founder_1_index"];
                    founder_1_region = generation_ooo_dictionary["founder_1_region"];
                    founder_1_district = generation_ooo_dictionary["founder_1_district"];
                    locality_founder_1 = generation_ooo_dictionary["locality_founder_1"];
                    founder_1_name_city = generation_ooo_dictionary["founder_1_name_city"];
                    founder_1_street = generation_ooo_dictionary["founder_1_street"];
                    founder_1_korpus = generation_ooo_dictionary["founder_1_korpus"];
                    founder_1_number_house = generation_ooo_dictionary["founder_1_number_house"];
                    founder_1_app = generation_ooo_dictionary["founder_1_app"];
                    founder_1_id_passport = generation_ooo_dictionary["founder_1_id_passport"];
                    founder_1_number_passport = generation_ooo_dictionary["founder_1_number_passport"];
                    founder_1_date_passport = generation_ooo_dictionary["founder_1_date_passport"];
                    founder_1_valid_passport = generation_ooo_dictionary["founder_1_valid_passport"];
                    founder_1_registration_passport = generation_ooo_dictionary["founder_1_registration_passport"];
                    founder_1_code = generation_ooo_dictionary["founder_1_code"];
                    founder_1_phone = generation_ooo_dictionary["founder_1_phone"];
                    founder_1_fond = generation_ooo_dictionary["founder_1_fond"];
                    founder_1_precent_fond = generation_ooo_dictionary["founder_1_precent_fond"];
                    founder_1_type_document = generation_ooo_dictionary["founder_1_type_document"];
                    founder_1_citizen = generation_ooo_dictionary["founder_1_citizen"];
                    founder_1_country = generation_ooo_dictionary["founder_1_country"];
                    switch (founder_1_type_document)
                    {
                        case "founder_1_passport":
                            founder_1_type_document = "Паспорт";
                            break;
                        case "founder_1_residence":
                            founder_1_type_document = "Вид на жительство";
                            break;
                    }
                    switch (founder_1_sex)
                    {
                        case "founder_1_sex_male":
                            founder_1_sex = "Мужской";
                            break;
                        case "founder_1_sex_female":
                            founder_1_sex = "Женский";
                            break;
                    }
                    founder_2_firstname = generation_ooo_dictionary["founder_2_firstname"];
                    founder_2_name = generation_ooo_dictionary["founder_2_name"];
                    founder_2_lastname = generation_ooo_dictionary["founder_2_lastname"];
                    founder_2_sex = generation_ooo_dictionary["founder_2_sex"];
                    founder_2_datebirthd = generation_ooo_dictionary["founder_2_datebirthd"];
                    founder_2_placebirthd = generation_ooo_dictionary["founder_2_placebirthd"];
                    founder_2_index = generation_ooo_dictionary["founder_2_index"];
                    founder_2_region = generation_ooo_dictionary["founder_2_region"];
                    founder_2_district = generation_ooo_dictionary["founder_2_district"];
                    locality_founder_2 = generation_ooo_dictionary["locality_founder_2"];
                    founder_2_name_city = generation_ooo_dictionary["founder_2_name_city"];
                    founder_2_street = generation_ooo_dictionary["founder_2_street"];
                    founder_2_korpus = generation_ooo_dictionary["founder_2_korpus"];
                    founder_2_number_house = generation_ooo_dictionary["founder_2_number_house"];
                    founder_2_app = generation_ooo_dictionary["founder_2_app"];
                    founder_2_id_passport = generation_ooo_dictionary["founder_2_id_passport"];
                    founder_2_number_passport = generation_ooo_dictionary["founder_2_number_passport"];
                    founder_2_date_passport = generation_ooo_dictionary["founder_2_date_passport"];
                    founder_2_valid_passport = generation_ooo_dictionary["founder_2_valid_passport"];
                    founder_2_registration_passport = generation_ooo_dictionary["founder_2_registration_passport"];
                    founder_2_code = generation_ooo_dictionary["founder_2_code"];
                    founder_2_phone = generation_ooo_dictionary["founder_2_phone"];
                    founder_2_fond = generation_ooo_dictionary["founder_2_fond"];
                    founder_2_precent_fond = generation_ooo_dictionary["founder_2_precent_fond"];
                    founder_2_type_document = generation_ooo_dictionary["founder_2_type_document"];
                    founder_2_citizen = generation_ooo_dictionary["founder_2_citizen"];
                    founder_2_country = generation_ooo_dictionary["founder_2_country"];
                    switch (founder_2_type_document)
                    {
                        case "founder_2_passport":
                            founder_2_type_document = "Паспорт";
                            break;
                        case "founder_2_residence":
                            founder_2_type_document = "Вид на жительство";
                            break;
                    }
                    switch (founder_2_sex)
                    {
                        case "founder_2_sex_male":
                            founder_2_sex = "Мужской";
                            break;
                        case "founder_2_sex_female":
                            founder_2_sex = "Женский";
                            break;
                    }
                    break;
                case 3:
                    founder_1_firstname = generation_ooo_dictionary["founder_1_firstname"];
                    founder_1_name = generation_ooo_dictionary["founder_1_name"];
                    founder_1_lastname = generation_ooo_dictionary["founder_1_lastname"];
                    founder_1_sex = generation_ooo_dictionary["founder_1_sex"];
                    founder_1_datebirthd = generation_ooo_dictionary["founder_1_datebirthd"];
                    founder_1_placebirthd = generation_ooo_dictionary["founder_1_placebirthd"];
                    founder_1_index = generation_ooo_dictionary["founder_1_index"];
                    founder_1_region = generation_ooo_dictionary["founder_1_region"];
                    founder_1_district = generation_ooo_dictionary["founder_1_district"];
                    locality_founder_1 = generation_ooo_dictionary["locality_founder_1"];
                    founder_1_name_city = generation_ooo_dictionary["founder_1_name_city"];
                    founder_1_street = generation_ooo_dictionary["founder_1_street"];
                    founder_1_korpus = generation_ooo_dictionary["founder_1_korpus"];
                    founder_1_number_house = generation_ooo_dictionary["founder_1_number_house"];
                    founder_1_app = generation_ooo_dictionary["founder_1_app"];
                    founder_1_id_passport = generation_ooo_dictionary["founder_1_id_passport"];
                    founder_1_number_passport = generation_ooo_dictionary["founder_1_number_passport"];
                    founder_1_date_passport = generation_ooo_dictionary["founder_1_date_passport"];
                    founder_1_valid_passport = generation_ooo_dictionary["founder_1_valid_passport"];
                    founder_1_registration_passport = generation_ooo_dictionary["founder_1_registration_passport"];
                    founder_1_code = generation_ooo_dictionary["founder_1_code"];
                    founder_1_phone = generation_ooo_dictionary["founder_1_phone"];
                    founder_1_fond = generation_ooo_dictionary["founder_1_fond"];
                    founder_1_precent_fond = generation_ooo_dictionary["founder_1_precent_fond"];
                    founder_1_type_document = generation_ooo_dictionary["founder_1_type_document"];
                    founder_1_citizen = generation_ooo_dictionary["founder_1_citizen"];
                    founder_1_country = generation_ooo_dictionary["founder_1_country"];
                    switch (founder_1_type_document)
                    {
                        case "founder_1_passport":
                            founder_1_type_document = "Паспорт";
                            break;
                        case "founder_1_residence":
                            founder_1_type_document = "Вид на жительство";
                            break;
                    }
                    switch (founder_1_sex)
                    {
                        case "founder_1_sex_male":
                            founder_1_sex = "Мужской";
                            break;
                        case "founder_1_sex_female":
                            founder_1_sex = "Женский";
                            break;
                    }
                    founder_2_firstname = generation_ooo_dictionary["founder_2_firstname"];
                    founder_2_name = generation_ooo_dictionary["founder_2_name"];
                    founder_2_lastname = generation_ooo_dictionary["founder_2_lastname"];
                    founder_2_sex = generation_ooo_dictionary["founder_2_sex"];
                    founder_2_datebirthd = generation_ooo_dictionary["founder_2_datebirthd"];
                    founder_2_placebirthd = generation_ooo_dictionary["founder_2_placebirthd"];
                    founder_2_index = generation_ooo_dictionary["founder_2_index"];
                    founder_2_region = generation_ooo_dictionary["founder_2_region"];
                    founder_2_district = generation_ooo_dictionary["founder_2_district"];
                    locality_founder_2 = generation_ooo_dictionary["locality_founder_2"];
                    founder_2_name_city = generation_ooo_dictionary["founder_2_name_city"];
                    founder_2_street = generation_ooo_dictionary["founder_2_street"];
                    founder_2_korpus = generation_ooo_dictionary["founder_2_korpus"];
                    founder_2_number_house = generation_ooo_dictionary["founder_2_number_house"];
                    founder_2_app = generation_ooo_dictionary["founder_2_app"];
                    founder_2_id_passport = generation_ooo_dictionary["founder_2_id_passport"];
                    founder_2_number_passport = generation_ooo_dictionary["founder_2_number_passport"];
                    founder_2_date_passport = generation_ooo_dictionary["founder_2_date_passport"];
                    founder_2_valid_passport = generation_ooo_dictionary["founder_2_valid_passport"];
                    founder_2_registration_passport = generation_ooo_dictionary["founder_2_registration_passport"];
                    founder_2_code = generation_ooo_dictionary["founder_2_code"];
                    founder_2_phone = generation_ooo_dictionary["founder_2_phone"];
                    founder_2_fond = generation_ooo_dictionary["founder_2_fond"];
                    founder_2_precent_fond = generation_ooo_dictionary["founder_2_precent_fond"];
                    founder_2_type_document = generation_ooo_dictionary["founder_2_type_document"];
                    founder_2_citizen = generation_ooo_dictionary["founder_2_citizen"];
                    founder_2_country = generation_ooo_dictionary["founder_2_country"];
                    switch (founder_2_type_document)
                    {
                        case "founder_2_passport":
                            founder_2_type_document = "Паспорт";
                            break;
                        case "founder_2_residence":
                            founder_2_type_document = "Вид на жительство";
                            break;
                    }
                    switch (founder_2_sex)
                    {
                        case "founder_2_sex_male":
                            founder_2_sex = "Мужской";
                            break;
                        case "founder_2_sex_female":
                            founder_2_sex = "Женский";
                            break;
                    }
                    
                    founder_3_firstname = generation_ooo_dictionary["founder_3_firstname"];
                    founder_3_name = generation_ooo_dictionary["founder_3_name"];
                    founder_3_lastname = generation_ooo_dictionary["founder_3_lastname"];
                    founder_3_sex = generation_ooo_dictionary["founder_3_sex"];
                    founder_3_datebirthd = generation_ooo_dictionary["founder_3_datebirthd"];
                    founder_3_placebirthd = generation_ooo_dictionary["founder_3_placebirthd"];
                    founder_3_index = generation_ooo_dictionary["founder_3_index"];
                    founder_3_region = generation_ooo_dictionary["founder_3_region"];
                    founder_3_district = generation_ooo_dictionary["founder_3_district"];
                    locality_founder_3 = generation_ooo_dictionary["locality_founder_3"];
                    founder_3_name_city = generation_ooo_dictionary["founder_3_name_city"];
                    founder_3_street = generation_ooo_dictionary["founder_3_street"];
                    founder_3_korpus = generation_ooo_dictionary["founder_3_korpus"];
                    founder_3_number_house = generation_ooo_dictionary["founder_3_number_house"];
                    founder_3_app = generation_ooo_dictionary["founder_3_app"];
                    founder_3_id_passport = generation_ooo_dictionary["founder_3_id_passport"];
                    founder_3_number_passport = generation_ooo_dictionary["founder_3_number_passport"];
                    founder_3_date_passport = generation_ooo_dictionary["founder_3_date_passport"];
                    founder_3_valid_passport = generation_ooo_dictionary["founder_3_valid_passport"];
                    founder_3_registration_passport = generation_ooo_dictionary["founder_3_registration_passport"];
                    founder_3_code = generation_ooo_dictionary["founder_3_code"];
                    founder_3_phone = generation_ooo_dictionary["founder_3_phone"];
                    founder_3_fond = generation_ooo_dictionary["founder_3_fond"];
                    founder_3_precent_fond = generation_ooo_dictionary["founder_3_precent_fond"];
                    founder_3_type_document = generation_ooo_dictionary["founder_3_type_document"];
                    founder_3_citizen = generation_ooo_dictionary["founder_3_citizen"];
                    founder_3_country = generation_ooo_dictionary["founder_3_country"];
                    switch (founder_3_type_document)
                    {
                        case "founder_3_passport":
                            founder_3_type_document = "Паспорт";
                            break;
                        case "founder_3_residence":
                            founder_3_type_document = "Вид на жительство";
                            break;
                    }

                    switch (founder_3_sex)
                    {
                        case "founder_3_sex_male":
                            founder_3_sex = "Мужской";
                            break;
                        case "founder_3_sex_female":
                            founder_3_sex = "Женский";
                            break;
                    }
                    break;

            }
            //Step 5

            // Step 6//



            generation_ooo_dictionary.Clear();
            generation_ooo_dictionary = null;




        }

        public void Generate()
        {



       switch (quantity_founder)
            {
                case 1:
                Docx file = new Docx("template_list_b.docx");

                file.SetValue("founder_1_type_document", founder_1_type_document);

                file.SetValue("founder_1_citizen", founder_1_citizen);

                file.SetValue("founder_1_country", founder_1_country);

                file.SetValue("director_type_document", director_type_document);

                file.SetValue("founder_1_firstname", founder_1_firstname);
                file.SetValue("founder_1_name", founder_1_name);
                file.SetValue("founder_1_lastname", founder_1_lastname);
                file.SetValue("founder_1_sex", founder_1_sex);
                file.SetValue("founder_1_datebirthd", founder_1_datebirthd);
                file.SetValue("founder_1_placebirthd", founder_1_placebirthd);
                file.SetValue("founder_1_index", founder_1_index);
                file.SetValue("founder_1_region", founder_1_region);
                file.SetValue("founder_1_district", founder_1_district);
                switch (locality_founder_1)
                {
                    case "city_founder_1":
                        file.SetValue("founder_1_city", founder_1_name_city);
                        file.SetValue("founder_1_minicity", "");
                        file.SetValue("founder_1_village", "");
                        break;
                    case "minicity_founder_1":
                        file.SetValue("founder_1_minicity", founder_1_name_city);
                        file.SetValue("founder_1_village", "");
                        file.SetValue("founder_1_city", "");

                        break;
                    case "village_founder_1":
                        file.SetValue("founder_1_village", founder_1_name_city);
                        file.SetValue("founder_1_city", "");
                        file.SetValue("founder_1_minicity", "");
                        break;
                }
                file.SetValue("founder_1_street", founder_1_street);
                file.SetValue("founder_1_korpus", founder_1_korpus);

                file.SetValue("founder_1_number_house", founder_1_number_house);
                file.SetValue("founder_1_app", founder_1_app);
                file.SetValue("founder_1_id_passport", founder_1_id_passport);
                file.SetValue("founder_1_number_passport", founder_1_number_passport);
                file.SetValue("founder_1_date_passport", founder_1_date_passport);
                file.SetValue("founder_1_valid_passport", founder_1_valid_passport);
                file.SetValue("founder_1_registration_passport", founder_1_registration_passport);

                file.SetValue("founder_1_code", founder_1_code);
                file.SetValue("founder_1_phone", founder_1_phone);

                file.SetValue("founder_1_fond", founder_1_fond);
                file.SetValue("founder_1_precent_fond", founder_1_precent_fond);
                file.SaveDocument(name_firma_ru + "_List_1.docx");
                    file.Dispose();
                  
              
                /// Reshenie ////

                file = new Docx("template_reshenie.docx");
                file.SetValue("firstname_director", firstname_director);
                file.SetValue("name_director", name_director);
                file.SetValue("lastname_director", lastname_director);
                file.SetValue("datebirdth_director", datebirthd_director);
                file.SetValue("placebirthd_director", placebirthd_director);
                file.SetValue("director_country", director_country);
                file.SetValue("index_director", index_director);
                file.SetValue("founder_1_type_document", founder_1_type_document);

                file.SetValue("founder_1_citizen", founder_1_citizen);

                file.SetValue("founder_1_country", founder_1_country);

                file.SetValue("director_type_document", director_type_document);
                if (region_director.ToLower() == "нет")
                {
                    file.SetValue("region_director", "");

                }
                else
                {
                    file.SetValue("region_director", region_director + " область,");
                }

                if (district_director.ToLower() == "нет")
                {

                    file.SetValue("district_director", "");
                }
                else
                {
                    file.SetValue("district_director", district_director + " район,");
                }
                file.SetValue("street_director", street_director);


                switch (locality_director)
                {
                    case "city_director":
                        file.SetValue("city_director", "город " + name_city_director + ", ");

                        break;
                    case "minicity_director":
                        file.SetValue("city_director", "посёлок " + name_city_director + ", ");


                        break;
                    case "village_director":
                        file.SetValue("city_director", "сельский совет " + name_city_director + ", ");

                        break;
                }

                    file.SetValue("house_director", ", д. " + house_director);
                    if (korpus_director.ToLower() == "нет")
                {
                    file.SetValue("korpus_director", "");

                }
                else
                {
                    file.SetValue("korpus_director", ", корпус: " + korpus_director);
                }
              

                if (app_director.ToLower() == "частный дом")
                {
                    file.SetValue("app_director", ".");
                }
                else
                {
                    file.SetValue("app_director", ", кв. " + app_director+".");

                }








                file.SetValue("idpassport_director", idpassport_director);
                file.SetValue("numberpassport_director", numberpassport_director);
                file.SetValue("datepassport_director", datapassport_director);
                file.SetValue("validpassport_director", validpassport_director);
                file.SetValue("registrationpassport_director", authority_passport_director);


                file.SetValue("name_firma_ru", name_firma_ru);

                file.SetValue("oked_name", oked_name);
                file.SetValue("oked", oked);


                file.SetValue("date_protocol", date_protocol);




                if (region_ur.ToLower() == "нет")
                {
                    file.SetValue("region_ur", "");

                }
                else
                {
                    file.SetValue("region_ur", region_ur + " область,");
                }

                if (district_ur.ToLower() == "нет")
                {

                    file.SetValue("district_ur", "");
                }
                else
                {
                    file.SetValue("district_ur", district_ur + " район,");
                }
                switch (locality_ur)
                {
                    case "city_ur":
                        file.SetValue("city_ur", "город " + name_city_ur + ", ");

                        break;
                    case "minicity_ur":
                        file.SetValue("city_ur", "посёлок " + name_city_ur + ", ");


                        break;
                    case "village_ur":
                        file.SetValue("city_ur", "сельский совет " + name_city_ur + ", ");

                        break;
                }
                file.SetValue("street_ur", street_ur);
                file.SetValue("house_ur", ", д. " + house_ur);

                if (korpus_ur.ToLower() == "нет")
                {

                    file.SetValue("korpus_ur", ", ");
                }
                else
                {
                    file.SetValue("korpus_ur", ", корпус: " + korpus_ur + ",");
                }
                file.SetValue("office_ur", office_ur);
                file.SetValue("founder_1_country", founder_1_country);

                file.SetValue("founder_1_firstname", founder_1_firstname);
                file.SetValue("founder_1_name", founder_1_name);
                file.SetValue("founder_1_lastname", founder_1_lastname);

                file.SetValue("founder_1_datebirdth", founder_1_datebirthd);
                file.SetValue("founder_1_index", founder_1_index);

                if (founder_1_region.ToLower() == "нет")
                {
                    file.SetValue("founder_1_region", "");

                }
                else
                {
                    file.SetValue("founder_1_region", founder_1_region + " область,");
                }

                if (founder_1_district.ToLower() == "нет")
                {

                    file.SetValue("founder_1_district", "");
                }
                else
                {
                    file.SetValue("founder_1_district", founder_1_district + " район,");
                }
                file.SetValue("founder_1_street", founder_1_street);
                switch (locality_founder_1)
                {
                    case "city_founder_1":
                        file.SetValue("founder_1_city", "город " + founder_1_name_city + ", ");

                        break;
                    case "minicity_founder_1":
                        file.SetValue("founder_1_city", "посёлок " + founder_1_name_city + ", ");


                        break;
                    case "village_founder_1":
                        file.SetValue("founder_1_city", "сельский совет " + founder_1_name_city + ", ");

                        break;
                }


                if (founder_1_korpus.ToLower() == "нет")
                {
                    file.SetValue("founder_1_korpus", "");

                }
                else
                {
                    file.SetValue("founder_1_korpus", ", корпус: " + founder_1_korpus + ",");
                }
                file.SetValue("founder_1_house", ", д. " + founder_1_number_house + ",");

                if (founder_1_app.ToLower() == "частный дом")
                {
                    file.SetValue("founder_1_app", "");
                }
                else
                {
                    file.SetValue("founder_1_app", "кв. " + founder_1_app);

                }
                file.SetValue("founder_1_idpassport", founder_1_id_passport);
                file.SetValue("founder_1_numberpassport", founder_1_number_passport);
                file.SetValue("founder_1_datepassport", founder_1_date_passport);
                file.SetValue("founder_1_registrationpassport", founder_1_registration_passport);
                file.SetValue("founder_1_fond", founder_1_fond);
                file.SaveDocument(name_firma_ru + "_Решение.docx");
                    file.Dispose();
                
             
                /*
               *  Устав
               */
                file = new Docx("templateustav_1founder_ooo.docx");
                file.SetValue("name_firma_ru", name_firma_ru);
                file.SetValue("date_protocol", date_protocol);

                file.SetValue("year", DateTime.Now.Year.ToString());
                file.SetValue("name_firma_by", name_firma_by);
                file.SetValue("founder_1_firstname", founder_1_firstname);
                file.SetValue("founder_1_name", founder_1_name);
                file.SetValue("founder_1_lastname", founder_1_lastname);
                file.SetValue("founder_1_datebirthd", founder_1_datebirthd);
                file.SetValue("founder_1_placebirthd", founder_1_placebirthd);
                file.SetValue("founder_1_index", founder_1_index);
                file.SetValue("founder_1_type_document", founder_1_type_document);

                file.SetValue("founder_1_citizen", founder_1_citizen);

                file.SetValue("founder_1_country", founder_1_country);

                file.SetValue("director_type_document", director_type_document);
                if (founder_1_region.ToLower() == "нет")
                {
                    file.SetValue("founder_1_region", "");

                }
                else
                {
                    file.SetValue("founder_1_region", "" + founder_1_region + " область,");
                }

                if (founder_1_district.ToLower() == "нет")
                {

                    file.SetValue("founder_1_district", "");
                }
                else
                {
                    file.SetValue("founder_1_district", "" + founder_1_district + " район,");
                }
                file.SetValue("founder_1_street", founder_1_street);
                switch (locality_founder_1)
                {
                    case "city_founder_1":
                        file.SetValue("founder_1_city", "город " + founder_1_name_city + ", ");

                        break;
                    case "minicity_founder_1":
                        file.SetValue("founder_1_city", "посёлок " + founder_1_name_city + ", ");


                        break;
                    case "village_founder_1":
                        file.SetValue("founder_1_city", "сельский совет " + founder_1_name_city + ", ");

                        break;
                }


                if (founder_1_korpus.ToLower() == "нет")
                {
                    file.SetValue("founder_1_korpus", "");

                }
                else
                {
                    file.SetValue("founder_1_korpus", ", корпус: " + founder_1_korpus + ",");
                }
                file.SetValue("founder_1_house", ", д. " + founder_1_number_house);

                if (founder_1_app.ToLower() == "частный дом")
                {
                    file.SetValue("founder_1_app", " ");
                }
                else
                {
                    file.SetValue("founder_1_app", ", кв. " + founder_1_app);

                }
                file.SetValue("founder_1_id_passport", founder_1_id_passport);
                file.SetValue("founder_1_number_passport", founder_1_number_passport);
                file.SetValue("founder_1_date_passport", founder_1_date_passport);
                file.SetValue("founder_1_valid_passport", founder_1_valid_passport);
                file.SetValue("founder_1_registration_passport", founder_1_registration_passport);
                file.SetValue("founder_1_fond", founder_1_fond);


                if (region_ur.ToLower() == "нет")
                {
                    file.SetValue("region_ur", "");

                }
                else
                {
                    file.SetValue("region_ur", "" + region_ur + " область,");
                }

                if (district_ur.ToLower() == "нет")
                {

                    file.SetValue("district_ur", "");
                }
                else
                {
                    file.SetValue("district_ur", "" + district_ur + " район,");
                }
                switch (locality_ur)
                {
                    case "city_ur":
                        file.SetValue("city_ur", "город " + name_city_ur + ", ");

                        break;
                    case "minicity_ur":
                        file.SetValue("city_ur", "посёлок " + name_city_ur + ", ");


                        break;
                    case "village_ur":
                        file.SetValue("city_ur", "сельский совет " + name_city_ur + ", ");

                        break;
                }
                file.SetValue("street_ur", street_ur);
                file.SetValue("house_ur", ", д. " + house_ur);

                if (korpus_ur.ToLower() == "нет")
                {

                    file.SetValue("korpus_ur", ", ");
                }
                else
                {
                    file.SetValue("korpus_ur", ", корпус: " + korpus_ur + ",");
                }
                file.SetValue("office_ur", office_ur);

                file.SaveDocument(name_firma_ru + "_Устав.docx");
               
                    file.Dispose();
                   
        
                /*
                  *  Заявление
                  */
                file = new Docx("template_1founder_gos.docx");

                file.SetValue("registration", registering_authority);
                file.SetValue("soglas", soglas);
                file.SetValue("name_firma_ru", name_firma_ru);
                file.SetValue("name_firma_by", name_firma_by);
                file.SetValue("date_protocol_2", date_protocol_2);
                file.SetValue("founder_1_type_document", founder_1_type_document);

                file.SetValue("founder_1_citizen", founder_1_citizen);

                file.SetValue("founder_1_country", founder_1_country);

                file.SetValue("director_type_document", director_type_document);
                file.SetValue("firstname_director", firstname_director);
                file.SetValue("name_director", name_director);
                file.SetValue("lastname_director", lastname_director);
                file.SetValue("datebirthd_director", datebirthd_director);
                file.SetValue("placebirthd_director", placebirthd_director);
                file.SetValue("director_country", director_country);
                file.SetValue("index_director", index_director);
                file.SetValue("region_director", region_director);
                file.SetValue("district_director", district_director);
                switch (locality_director)
                {
                    case "city_director":
                        file.SetValue("city_director", name_city_director);
                        file.SetValue("minicity_director", "");
                        file.SetValue("village_director", "");
                        break;
                    case "minicity_director":
                        file.SetValue("minicity_director", name_city_director);
                        file.SetValue("city_director", "");
                        file.SetValue("village_director", "");
                        break;
                    case "village_director":
                        file.SetValue("village_director", name_city_director);
                        file.SetValue("minicity_director", "");
                        file.SetValue("city_director", "");
                        break;
                }

                file.SetValue("street_director", street_director);

                file.SetValue("korpus_director", korpus_director);

                file.SetValue("house_director", house_director);
                file.SetValue("app_director", app_director);
                file.SetValue("idpassport_director", idpassport_director);
                file.SetValue("numberpassport_director", numberpassport_director);
                file.SetValue("datepassport_director", datapassport_director);
                file.SetValue("validpassport_director", validpassport_director);
                file.SetValue("registrationpassport_director", authority_passport_director);
                file.SetValue("code_director", code_director);
                file.SetValue("phone_director", phone_director);
                file.SetValue("oked_name", oked_name);
                file.SetValue("oked", oked);
                file.SetValue("all_fond", all_fond);
                file.SetValue("region_ur", region_ur);
                file.SetValue("district_ur", district_ur);
                switch (locality_ur)
                {
                    case "city_ur":
                        file.SetValue("city_ur", name_city_ur);
                        file.SetValue("minicity_ur", "");
                        file.SetValue("village_ur", "");
                        break;
                    case "minicity_ur":
                        file.SetValue("minicity_ur", name_city_ur);
                        file.SetValue("city_ur", "");
                        file.SetValue("village_ur", "");
                        break;
                    case "village_ur":
                        file.SetValue("village_ur", name_city_ur);
                        file.SetValue("minicity_ur", "");
                        file.SetValue("city_ur", "");
                        break;
                }
                file.SetValue("street_ur", street_ur);
                file.SetValue("index_ur", index_ur);
                file.SetValue("house_ur", house_ur);
                file.SetValue("korpus_ur", korpus_ur);
                file.SetValue("office_ur", office_ur);
                file.SaveDocument(name_firma_ru + "_Registration.docx"); //
                    file.Dispose();


                    MessageBox.Show("Завершено!");
               
                    break;

                case 2:

                 file = new Docx("template_list_b.docx");
                file.SetValue("founder_1_type_document", founder_1_type_document);
                file.SetValue("founder_2_type_document", founder_2_type_document);
                file.SetValue("founder_1_citizen", founder_1_citizen);
                file.SetValue("founder_2_citizen", founder_2_citizen);
                file.SetValue("founder_1_country", founder_1_country);
                file.SetValue("founder_2_country", founder_2_country);

                file.SetValue("founder_1_firstname", founder_1_firstname);
                file.SetValue("founder_1_name", founder_1_name);
                file.SetValue("founder_1_lastname", founder_1_lastname);
                file.SetValue("founder_1_sex", founder_1_sex);
                file.SetValue("founder_1_datebirthd", founder_1_datebirthd);
                file.SetValue("founder_1_placebirthd", founder_1_placebirthd);
                file.SetValue("founder_1_index", founder_1_index);
                file.SetValue("founder_1_region", founder_1_region);
                file.SetValue("founder_1_district", this.founder_1_district);
                switch (locality_founder_1)
                {
                    case "city_founder_1":
                        file.SetValue("founder_1_city", founder_1_name_city);
                        file.SetValue("founder_1_minicity", "");
                        file.SetValue("founder_1_village", "");
                        break;
                    case "minicity_founder_1":
                        file.SetValue("founder_1_minicity", founder_1_name_city);
                        file.SetValue("founder_1_village", "");
                        file.SetValue("founder_1_city", "");

                        break;
                    case "village_founder_1":
                        file.SetValue("founder_1_village", founder_1_name_city);
                        file.SetValue("founder_1_city", "");
                        file.SetValue("founder_1_minicity", "");
                        break;
                }
                file.SetValue("founder_1_street", founder_1_street);
                file.SetValue("founder_1_korpus", founder_1_korpus);

                file.SetValue("founder_1_number_house", founder_1_number_house);
                file.SetValue("founder_1_app", founder_1_app);
                file.SetValue("founder_1_id_passport", founder_1_id_passport);
                file.SetValue("founder_1_number_passport", founder_1_number_passport);
                file.SetValue("founder_1_date_passport", founder_1_date_passport);
                file.SetValue("founder_1_valid_passport", founder_1_valid_passport);
                file.SetValue("founder_1_registration_passport", founder_1_registration_passport);

                file.SetValue("founder_1_code", founder_1_code);
                file.SetValue("founder_1_phone", founder_1_phone);

                file.SetValue("founder_1_fond", founder_1_fond);
                file.SetValue("founder_1_precent_fond", founder_1_precent_fond);


                file.SaveDocument(name_firma_ru + "_List_1.docx");
            
                    file.Dispose();
                    
                file = new Docx("template_list_a.docx");
                file.SetValue("founder_1_type_document", founder_1_type_document);
                file.SetValue("founder_2_type_document", founder_2_type_document);
                file.SetValue("founder_1_citizen", founder_1_citizen);
                file.SetValue("founder_2_citizen", founder_2_citizen);
                file.SetValue("founder_1_country", founder_1_country);
                file.SetValue("founder_2_country", founder_2_country);
                file.SetValue("founder_2_firstname", founder_2_firstname);
                file.SetValue("founder_2_name", founder_2_name);
                file.SetValue("founder_2_lastname", founder_2_lastname);
                file.SetValue("founder_2_sex", founder_2_sex);
                file.SetValue("founder_2_datebirthd", founder_2_datebirthd);
                file.SetValue("founder_2_placebirthd", founder_2_placebirthd);
                file.SetValue("founder_2_index", founder_2_index);
                file.SetValue("founder_2_region", founder_2_region);
                file.SetValue("founder_2_district", founder_2_district);
                switch (locality_founder_2)
                {
                    case "city_founder_2":
                        file.SetValue("founder_2_city", founder_2_name_city);
                        file.SetValue("founder_2_minicity", "");
                        file.SetValue("founder_2_village", "");
                        break;
                    case "minicity_founder_2":
                        file.SetValue("founder_2_minicity", founder_2_name_city);
                        file.SetValue("founder_2_village", "");
                        file.SetValue("founder_2_city", "");

                        break;
                    case "village_founder_2":
                        file.SetValue("founder_2_village", founder_2_name_city);
                        file.SetValue("founder_2_city", "");
                        file.SetValue("founder_2_minicity", "");
                        break;
                }
                file.SetValue("founder_2_street", founder_2_street);
                file.SetValue("founder_2_korpus", founder_2_korpus);

                file.SetValue("founder_2_number_house", founder_2_number_house);
                file.SetValue("founder_2_app", founder_2_app);
                file.SetValue("founder_2_id_passport", founder_2_id_passport);
                file.SetValue("founder_2_number_passport", founder_2_number_passport);
                file.SetValue("founder_2_date_passport", founder_2_date_passport);
                file.SetValue("founder_2_valid_passport", founder_2_valid_passport);
                file.SetValue("founder_2_registration_passport", founder_2_registration_passport);

                file.SetValue("founder_2_code", founder_2_code);
                file.SetValue("founder_2_phone", founder_2_phone);

                file.SetValue("founder_2_fond", founder_2_fond);
                file.SetValue("founder_2_precent_fond", founder_2_precent_fond);
                file.SaveDocument(name_firma_ru + "_List2.docx"); //
            
                    file.Dispose();
                   
                file = new Docx("template_protocol_1.docx");
                file.SetValue("name_firma_ru", name_firma_ru);

                file.SetValue("oked_name", oked_name);
                file.SetValue("oked", oked);

                file.SetValue("all_fond", all_fond);


                file.SetValue("date_protocol", date_protocol);
                file.SetValue("date_protocol_2", date_protocol_2);


                if (region_ur.ToLower() == "нет")
                {
                    file.SetValue("region_ur", " ");

                }
                else
                {
                    file.SetValue("region_ur", " " + region_ur + " область,");
                }

                if (district_ur.ToLower() == "нет")
                {

                    file.SetValue("district_ur", " ");
                }
                else
                {
                    file.SetValue("district_ur", " " + district_ur + " район,");
                }
                switch (locality_ur)
                {
                    case "city_ur":
                        file.SetValue("city_ur", "город " + name_city_ur + ",");

                        break;
                    case "minicity_ur":
                        file.SetValue("city_ur", "посёлок " + name_city_ur + ",");


                        break;
                    case "village_ur":
                        file.SetValue("city_ur", "сельский совет " + name_city_ur + ",");

                        break;
                }
                file.SetValue("street_ur", street_ur+", ");
                file.SetValue("house_ur", house_ur+", ");

                if (korpus_ur.ToLower() == "нет")
                {

                    file.SetValue("korpus_ur", "");
                }
                else
                {
                    file.SetValue("korpus_ur", "корпус: " + korpus_ur + ",");
                }
                file.SetValue("office_ur", office_ur);
                file.SetValue("founder_1_type_document", founder_1_type_document);
                file.SetValue("founder_2_type_document", founder_2_type_document);
                file.SetValue("founder_1_citizen", founder_1_citizen);
                file.SetValue("founder_2_citizen", founder_2_citizen);
                file.SetValue("founder_1_country", founder_1_country);
                file.SetValue("founder_2_country", founder_2_country);
                file.SetValue("founder_1_firstname", founder_1_firstname);
                file.SetValue("founder_1_name", founder_1_name);
                file.SetValue("founder_1_lastname", founder_1_lastname);

                file.SetValue("founder_1_datebirdth", founder_1_datebirthd);
                file.SetValue("founder_1_index", founder_1_index);



                if (founder_1_region.ToLower() == "нет")
                {
                    file.SetValue("founder_1_region", " ");

                }
                else
                {
                    file.SetValue("founder_1_region", founder_1_region + " область,");
                }

                if (this.founder_1_district.ToLower() == "нет")
                {

                    file.SetValue("founder_1_district", " ");
                }
                else
                {
                    file.SetValue("founder_1_district", this.founder_1_district + " район,");
                }
                file.SetValue("founder_1_street", founder_1_street + ", ");
                switch (locality_founder_1)
                {
                    case "city_founder_1":
                        file.SetValue("founder_1_city", "город " + founder_1_name_city + ",");

                        break;
                    case "minicity_founder_1":
                        file.SetValue("founder_1_city", "посёлок " + founder_1_name_city + ",");


                        break;
                    case "village_founder_1":
                        file.SetValue("founder_1_city", "сельский совет " + founder_1_name_city + ",");

                        break;
                }


                if (founder_1_korpus.ToLower() == "нет")
                {
                    file.SetValue("founder_1_korpus", "");

                }
                else
                {
                    file.SetValue("founder_1_korpus", " корпус: " + founder_1_korpus + ",");
                }
                file.SetValue("founder_1_house", " д. " + founder_1_number_house);

                if (founder_1_app.ToLower() == "частный дом")
                {
                    file.SetValue("founder_1_app", "");
                }
                else
                {
                    file.SetValue("founder_1_app", ", кв. " + founder_1_app);

                }
                file.SetValue("founder_1_idpassport", founder_1_id_passport);
                file.SetValue("founder_1_numberpassport", founder_1_number_passport);
                file.SetValue("founder_1_datepassport", founder_1_date_passport);
                file.SetValue("founder_1_registrationpassport", founder_1_registration_passport);

                file.SetValue("founder_1_fond", founder_1_fond);
                file.SetValue("founder_1_precent_fond", founder_1_precent_fond);

                file.SetValue("founder_2_firstname", founder_2_firstname);
                file.SetValue("founder_2_name", founder_2_name);
                file.SetValue("founder_2_lastname", founder_2_lastname);

                file.SetValue("founder_2_datebirdth", founder_2_datebirthd);
                file.SetValue("founder_2_index", founder_2_index);


                file.SetValue("founder_2_idpassport", founder_2_id_passport);
                file.SetValue("founder_2_numberpassport", founder_2_number_passport);
                file.SetValue("founder_2_datepassport", founder_2_date_passport);
                file.SetValue("founder_2_registrationpassport", founder_2_registration_passport);
                file.SetValue("founder_2_fond", founder_2_fond);
                file.SetValue("founder_2_precent_fond", founder_2_precent_fond);


                if (founder_2_region.ToLower() == "нет")
                {
                    file.SetValue("founder_2_region", "");

                }
                else
                {
                    file.SetValue("founder_2_region", founder_2_region + " область,");
                }

                if (founder_2_district.ToLower() == "нет")
                {

                    file.SetValue("founder_2_district", "");
                }
                else
                {
                    file.SetValue("founder_2_district", founder_2_district + " район,");
                }
                file.SetValue("founder_2_street", founder_2_street + ", ");
                switch (locality_founder_2)
                {
                    case "city_founder_2":
                        file.SetValue("founder_2_city", "город " + founder_2_name_city + ",");

                        break;
                    case "minicity_founder_2":
                        file.SetValue("founder_2_city", "посёлок " + founder_2_name_city + ",");


                        break;
                    case "village_founder_2":
                        file.SetValue("founder_2_city", "сельский совет " + founder_2_name_city + ",");

                        break;
                }


                if (founder_2_korpus.ToLower() == "нет")
                {
                    file.SetValue("founder_2_korpus", "");

                }
                else
                {
                    file.SetValue("founder_2_korpus", " корпус: " + founder_2_korpus + ",");
                }
                file.SetValue("founder_2_house", " д. " + founder_2_number_house);

                if (founder_2_app.ToLower() == "частный дом")
                {
                    file.SetValue("founder_2_app", "");
                }
                else
                {
                    file.SetValue("founder_2_app", ", кв." + founder_2_app);
                }

                file.SaveDocument(name_firma_ru + "_Протокол№1.docx");
                    file.Dispose();
                   
                    
                
                file = new Docx("template_protocol_2.docx");
                file.SetValue("founder_1_type_document", founder_1_type_document);
                file.SetValue("founder_2_type_document", founder_2_type_document);
                file.SetValue("founder_1_citizen", founder_1_citizen);
                file.SetValue("founder_2_citizen", founder_2_citizen);
                file.SetValue("founder_1_country", founder_1_country+", ");
                file.SetValue("founder_2_country", founder_2_country+", ");
                file.SetValue("director_type_document", director_type_document);

                file.SetValue("firstname_director", firstname_director);
                file.SetValue("name_director", name_director);
                file.SetValue("lastname_director", lastname_director);
                file.SetValue("datebirdth_director", datebirthd_director);
                file.SetValue("placebirthd_director", placebirthd_director);
                file.SetValue("index_director", index_director);
                file.SetValue("director_country", director_country);
                if (region_director.ToLower() == "нет") {
                    file.SetValue("region_director", "");

                }
        else {
                    file.SetValue("region_director", region_director+" область,");
                }

                if (district_director.ToLower() == "нет") {

                    file.SetValue("district_director", "");
                }
        else {
                    file.SetValue("district_director", district_director+" район,");
                }
                file.SetValue("street_director", ", "+street_director);


                switch (locality_director)
                {
                    case "city_director":
                        file.SetValue("city_director", "город "+name_city_director);

                        break;
                    case "minicity_director":
                        file.SetValue("city_director", "посёлок "+name_city_director);


                        break;
                    case "village_director":
                        file.SetValue("city_director", "сельский совет "+name_city_director);

                        break;
                }


                if (korpus_director.ToLower() == "нет" ) {
                    file.SetValue("korpus_director", "");

                }
        else {
                    file.SetValue("korpus_director", ", корпус: "+korpus_director+",");
                }
                file.SetValue("house_director", ", д. "+house_director);

                if (app_director.ToLower() == "частный дом")
                {
                    file.SetValue("app_director", "");
                }
                else
                {
                    file.SetValue("app_director", ", кв. "+app_director);

                }
                file.SetValue("idpassport_director", idpassport_director);
                file.SetValue("numberpassport_director", numberpassport_director);
                file.SetValue("datepassport_director", datapassport_director);
                file.SetValue("validpassport_director", validpassport_director);
                file.SetValue("registrationpassport_director", authority_passport_director);


                file.SetValue("name_firma_ru", name_firma_ru);

                file.SetValue("oked_name", oked_name);
                file.SetValue("oked", oked);


                file.SetValue("date_protocol", date_protocol);
                file.SetValue("date_protocol_2", date_protocol_2);




                if (region_ur.ToLower() == "нет") {
                    file.SetValue("region_ur", ", ");

                }
        else {
                    file.SetValue("region_ur", region_ur+" область,");
                }

                if (district_ur.ToLower() == "нет") {

                    file.SetValue("district_ur", " ");
                }
        else {
                    file.SetValue("district_ur", district_ur+" район,");
                }
                switch (locality_ur)
                {
                    case "city_ur":
                        file.SetValue("city_ur", "город "+name_city_ur);

                        break;
                    case "minicity_ur":
                        file.SetValue("city_ur", "посёлок "+name_city_ur);


                        break;
                    case "village_ur":
                        file.SetValue("city_ur", "сельский совет "+name_city_ur);

                        break;
                }
                file.SetValue("office_ur", office_ur);
                file.SetValue("founder_1_firstname", founder_1_firstname);
                file.SetValue("founder_1_name", founder_1_name);
                file.SetValue("founder_1_lastname", founder_1_lastname);

                file.SetValue("founder_1_datebirthd", founder_1_datebirthd);
                file.SetValue("founder_1_index", founder_1_index);

                if (founder_1_region.ToLower() == "нет") {
                    file.SetValue("founder_1_region", "");

                }
        else {
                    file.SetValue("founder_1_region", founder_1_region+" область, ");
                }

                if (this.founder_1_district.ToLower() == "нет") {

                    file.SetValue("founder_1_district", "");
                }
        else {
                    file.SetValue("founder_1_district", this.founder_1_district+" район,");
                }
                file.SetValue("founder_1_street", founder_1_street+", ");
                switch (locality_founder_1)
                {
                    case "city_founder_1":
                        file.SetValue("founder_1_city", " город "+founder_1_name_city+", ");

                        break;
                    case "minicity_founder_1":
                        file.SetValue("founder_1_city", " посёлок "+founder_1_name_city+",");


                        break;
                    case "village_founder_1":
                        file.SetValue("founder_1_city", " сельский совет "+founder_1_name_city+",");

                        break;
                }


                if (founder_1_korpus.ToLower() == "нет" ) {
                    file.SetValue("founder_1_korpus", " ");

                }
        else {
                    file.SetValue("founder_1_korpus", " корпус: "+founder_1_korpus+", ");
                }
                file.SetValue("founder_1_house", " д. "+founder_1_number_house+"");
                if (founder_1_app == "частный дом")
                {
                    file.SetValue("founder_1_app", "");
                }
                else
                {
                    file.SetValue("founder_1_app", ", кв. "+founder_1_app);

                }


                file.SetValue("founder_1_idpassport", founder_1_id_passport);
                file.SetValue("founder_1_numberpassport", founder_1_number_passport);
                file.SetValue("founder_1_datepassport", founder_1_date_passport);
                file.SetValue("founder_1_registrationpassport", founder_1_registration_passport);


                file.SetValue("founder_2_firstname", founder_2_firstname);
                file.SetValue("founder_2_name", founder_2_name);
                file.SetValue("founder_2_lastname", founder_2_lastname);

                file.SetValue("founder_2_datebirthd", founder_2_datebirthd);
                file.SetValue("founder_2_index", founder_2_index);




                if (founder_2_region.ToLower() == "нет") {
                    file.SetValue("founder_2_region", "");

                }
        else {
                    file.SetValue("founder_2_region", founder_2_region+" область, ");
                }

                if (founder_2_district.ToLower() == "нет") {

                    file.SetValue("founder_2_district", "");
                }
        else {
                    file.SetValue("founder_2_district", founder_2_district+" район,");
                }
                file.SetValue("founder_2_street", founder_2_street+", ");
                switch (locality_founder_2)
                {
                    case "city_founder_2":
                        file.SetValue("founder_2_city", " город "+founder_2_name_city+", ");

                        break;
                    case "minicity_founder_2":
                        file.SetValue("founder_2_city", " посёлок "+founder_2_name_city+",");


                        break;
                    case "village_founder_2":
                        file.SetValue("founder_2_city", " сельский совет "+founder_2_name_city+",");

                        break;
                }


                if (founder_2_korpus.ToLower() == "нет" ) {
                    file.SetValue("founder_2_korpus", " ");

                }
        else {
                    file.SetValue("founder_2_korpus", " корпус: "+founder_2_korpus+", ");
                }
                file.SetValue("founder_2_house", " д. "+founder_2_number_house+"");
                if (founder_2_app == "частный дом")
                {
                    file.SetValue("founder_2_app", "");
                }
                else
                {
                    file.SetValue("founder_2_app", ", кв. "+founder_2_app);

                }
                file.SetValue("founder_2_idpassport", founder_2_id_passport);
                file.SetValue("founder_2_numberpassport", founder_2_number_passport);
                file.SetValue("founder_2_datepassport", founder_2_date_passport);
                file.SetValue("founder_2_registrationpassport", founder_2_registration_passport);
             
                
              

                file.SaveDocument(name_firma_ru + "_Протокол№2.docx");
                  
                    file.Dispose();
                 
               
                file = new Docx("templateustav_ooo.docx");
                file.SetValue("name_firma_ru", name_firma_ru);
                file.SetValue("date_protocol", date_protocol);
                file.SetValue("year", DateTime.Now.Year.ToString());
                file.SetValue("name_firma_by", name_firma_by);
                file.SetValue("founder_1_firstname", founder_1_firstname);
                file.SetValue("founder_1_name", founder_1_name);
                file.SetValue("founder_1_lastname", founder_1_lastname);
                file.SetValue("founder_1_datebirthd", founder_1_datebirthd);
                file.SetValue("founder_1_placebirthd", founder_1_placebirthd);
                file.SetValue("founder_1_index", founder_1_index);
                file.SetValue("founder_1_type_document", founder_1_type_document);
                file.SetValue("founder_2_type_document", founder_2_type_document);
                file.SetValue("founder_1_citizen", founder_1_citizen);
                file.SetValue("founder_2_citizen", founder_2_citizen);
                file.SetValue("founder_1_country", founder_1_country);
                file.SetValue("founder_2_country", founder_2_country);


                if (founder_1_region.ToLower() == "нет" ) {
                    file.SetValue("founder_1_region", "");

                }
        else {
                    file.SetValue("founder_1_region", founder_1_region + " область,");
                }

                if (founder_1_district.ToLower() == "нет") {

                    file.SetValue("founder_1_district", "");
                }
        else {
                    file.SetValue("founder_1_district", founder_1_district + " район,");
                }
                file.SetValue("founder_1_street", founder_1_street + ", ");
                switch (locality_founder_1)
                {
                    case "city_founder_1":
                        file.SetValue("founder_1_city", "город "+founder_1_name_city + ", ");

                        break;
                    case "minicity_founder_1":
                        file.SetValue("founder_1_city", "посёлок "+founder_1_name_city + ", ");


                        break;
                    case "village_founder_1":
                        file.SetValue("founder_1_city", "сельский совет "+founder_1_name_city + ", ");

                        break;
                }


                if (founder_1_korpus.ToLower() == "нет") {
                    file.SetValue("founder_1_korpus", "");

                }
        else {
                    file.SetValue("founder_1_korpus", "корпус: " + founder_1_korpus + ", ");
                }
                file.SetValue("founder_1_house", " д. " + founder_1_number_house);

                if (founder_1_app == "частный дом")
                {
                    file.SetValue("founder_1_app", "");
                }
                else
                {
                    file.SetValue("founder_1_app", ", кв. "+founder_1_app);

                }
                file.SetValue("founder_1_id_passport", founder_1_id_passport);
                file.SetValue("founder_1_number_passport", founder_1_number_passport);
                file.SetValue("founder_1_date_passport", founder_1_date_passport);
                file.SetValue("founder_1_valid_passport", founder_1_valid_passport);
                file.SetValue("founder_1_registration_passport", founder_1_registration_passport);
                file.SetValue("founder_1_fond", founder_1_fond);
                file.SetValue("founder_1_precent_fond", founder_1_precent_fond);


                file.SetValue("founder_2_firstname", founder_2_firstname);
                file.SetValue("founder_2_name", founder_2_name);
                file.SetValue("founder_2_lastname", founder_2_lastname);
                file.SetValue("founder_2_datebirthd", founder_2_datebirthd);
                file.SetValue("founder_2_placebirthd", founder_2_placebirthd);
                file.SetValue("founder_2_index", founder_2_index);



                if (founder_2_region.ToLower() == "нет" ) {
                    file.SetValue("founder_2_region", "");

                }
        else {
                    file.SetValue("founder_2_region", founder_2_region + " область, ");
                }

                if (founder_2_district.ToLower() == "нет") {

                    file.SetValue("founder_2_district", "");
                }
        else {
                    file.SetValue("founder_2_district", founder_2_district + " район,");
                }
                file.SetValue("founder_2_street", founder_2_street + ", ");
                switch (locality_founder_2)
                {
                    case "city_founder_2":
                        file.SetValue("founder_2_city", " город "+founder_2_name_city+ ", ");

                        break;
                    case "minicity_founder_2":
                        file.SetValue("founder_2_city", " посёлок "+founder_2_name_city + ",");


                        break;
                    case "village_founder_2":
                        file.SetValue("founder_2_city", " сельский совет "+founder_2_name_city + ",");

                        break;
                }


                if (founder_2_korpus.ToLower() == "нет" ) {
                    file.SetValue("founder_2_korpus", " ");

                }
        else {
                    file.SetValue("founder_2_korpus", " корпус: " + founder_2_korpus + ", ");
                }
                file.SetValue("founder_2_house", " д. "+founder_2_number_house);
                if (founder_2_app == "частный дом")
                {
                    file.SetValue("founder_2_app", "");
                }
                else
                {
                    file.SetValue("founder_2_app", ", кв. "+founder_2_app);

                }
                file.SetValue("founder_2_id_passport", founder_2_id_passport);
                file.SetValue("founder_2_number_passport", founder_2_number_passport);
                file.SetValue("founder_2_date_passport", founder_2_date_passport);
                file.SetValue("founder_2_valid_passport", founder_2_valid_passport);
                file.SetValue("founder_2_registration_passport", founder_2_registration_passport);
                file.SetValue("founder_2_fond", founder_2_fond);
                file.SetValue("founder_2_precent_fond", founder_2_precent_fond);
                file.SetValue("all_fond", all_fond);


                if (region_ur.ToLower() == "нет" ) {
                    file.SetValue("region_ur", "");

                }
        else {
                    file.SetValue("region_ur", region_ur + " область,");
                }

                if (district_ur.ToLower() == "нет") {

                    file.SetValue("district_ur", "");
                }
        else {
                    file.SetValue("district_ur", district_ur+ " район,");
                }
                switch (locality_ur)
                {
                    case "city_ur":
                        file.SetValue("city_ur", " город "+ name_city_ur+ ",");

                        break;
                    case "minicity_ur":
                        file.SetValue("city_ur", " посёлок "+name_city_ur+ ",");


                        break;
                    case "village_ur":
                        file.SetValue("city_ur", " сельский совет "+name_city_ur+ ",");

                        break;
                }
                file.SetValue("street_ur", street_ur + ", ");
                file.SetValue("house_ur", " д. " +house_ur+ ", ");

                if (korpus_ur.ToLower() == "нет" ) {

                    file.SetValue("korpus_ur", " ");
                }
        else {
                    file.SetValue("korpus_ur", " корпус: " +korpus_ur+ ", ");
                }
                file.SetValue("office_ur", office_ur);
     file.SaveDocument(name_firma_ru+"_Устав.docx");//
               
                    file.Dispose();
                
                file = new Docx("template_gos.docx");
                file.SetValue("registration", registering_authority);
                file.SetValue("soglas", soglas);
                file.SetValue("name_firma_ru", name_firma_ru);
                file.SetValue("name_firma_by", name_firma_by);
                file.SetValue("date_protocol_2", date_protocol_2);

                file.SetValue("firstname_director", firstname_director);
                file.SetValue("name_director", name_director);
                file.SetValue("lastname_director", lastname_director);
                file.SetValue("datebirthd_director", datebirthd_director);
                file.SetValue("placebirthd_director", placebirthd_director);
                file.SetValue("director_country", director_country);
                file.SetValue("index_director", index_director);
                file.SetValue("region_director", region_director);
                file.SetValue("district_director", district_director);
                switch (locality_director)
                {
                    case "city_director":
                        file.SetValue("city_director", name_city_director);
                        file.SetValue("minicity_director", "");
                        file.SetValue("village_director", "");
                        break;
                    case "minicity_director":
                        file.SetValue("minicity_director", name_city_director);
                        file.SetValue("city_director", "");
                        file.SetValue("village_director", "");
                        break;
                    case "village_director":
                        file.SetValue("village_director", name_city_director);
                        file.SetValue("minicity_director", "");
                        file.SetValue("city_director", "");
                        break;
                }

                file.SetValue("street_director", street_director);

                file.SetValue("korpus_director", korpus_director);

                file.SetValue("house_director", house_director);
                file.SetValue("app_director", app_director);
                file.SetValue("idpassport_director", idpassport_director);
                file.SetValue("numberpassport_director", numberpassport_director);
                file.SetValue("datepassport_director", datapassport_director);
                file.SetValue("validpassport_director", validpassport_director);
                file.SetValue("registrationpassport_director", authority_passport_director);
                file.SetValue("code_director", code_director);
                file.SetValue("phone_director", phone_director);
                file.SetValue("oked_name", oked_name);
                file.SetValue("oked", oked);
                file.SetValue("all_fond", all_fond);
                file.SetValue("region_ur", region_ur);
                file.SetValue("district_ur", district_ur);
                switch (locality_ur)
                {
                    case "city_ur":
                        file.SetValue("city_ur", name_city_ur);
                        file.SetValue("minicity_ur", "");
                        file.SetValue("village_ur", "");
                        break;
                    case "minicity_ur":
                        file.SetValue("minicity_ur", name_city_ur);
                        file.SetValue("city_ur", "");
                        file.SetValue("village_ur", "");
                        break;
                    case "village_ur":
                        file.SetValue("village_ur", name_city_ur);
                        file.SetValue("minicity_ur", "");
                        file.SetValue("city_ur", "");
                        break;
                }
                file.SetValue("street_ur", street_ur);
                file.SetValue("index_ur", index_ur);
                file.SetValue("house_ur", house_ur);
                file.SetValue("korpus_ur", korpus_ur);
                file.SetValue("office_ur", office_ur);


                file.SetValue("director_type_document", director_type_document);
                file.SaveDocument(name_firma_ru+"_Регистрация.docx");
             
                    file.Dispose();
                    

                    MessageBox.Show("Завершено!");

                    break;
                case 3:
                    file = new Docx("template_list_b.docx");
                    file.SetValue("founder_1_type_document", founder_1_type_document);
                    file.SetValue("founder_1_citizen", founder_1_citizen);
                    file.SetValue("founder_1_country", founder_1_country);

                    file.SetValue("founder_1_firstname", founder_1_firstname);
                    file.SetValue("founder_1_name", founder_1_name);
                    file.SetValue("founder_1_lastname", founder_1_lastname);
                    file.SetValue("founder_1_sex", founder_1_sex);
                    file.SetValue("founder_1_datebirthd", founder_1_datebirthd);
                    file.SetValue("founder_1_placebirthd", founder_1_placebirthd);
                    file.SetValue("founder_1_index", founder_1_index);
                    file.SetValue("founder_1_region", founder_1_region);
                    file.SetValue("founder_1_district", this.founder_1_district);
                    switch (locality_founder_1)
                    {
                        case "city_founder_1":
                            file.SetValue("founder_1_city", founder_1_name_city);
                            file.SetValue("founder_1_minicity", "");
                            file.SetValue("founder_1_village", "");
                            break;
                        case "minicity_founder_1":
                            file.SetValue("founder_1_minicity", founder_1_name_city);
                            file.SetValue("founder_1_village", "");
                            file.SetValue("founder_1_city", "");

                            break;
                        case "village_founder_1":
                            file.SetValue("founder_1_village", founder_1_name_city);
                            file.SetValue("founder_1_city", "");
                            file.SetValue("founder_1_minicity", "");
                            break;
                    }
                    file.SetValue("founder_1_street", founder_1_street);
                    file.SetValue("founder_1_korpus", founder_1_korpus);

                    file.SetValue("founder_1_number_house", founder_1_number_house);
                    file.SetValue("founder_1_app", founder_1_app);
                    file.SetValue("founder_1_id_passport", founder_1_id_passport);
                    file.SetValue("founder_1_number_passport", founder_1_number_passport);
                    file.SetValue("founder_1_date_passport", founder_1_date_passport);
                    file.SetValue("founder_1_valid_passport", founder_1_valid_passport);
                    file.SetValue("founder_1_registration_passport", founder_1_registration_passport);

                    file.SetValue("founder_1_code", founder_1_code);
                    file.SetValue("founder_1_phone", founder_1_phone);

                    file.SetValue("founder_1_fond", founder_1_fond);
                    file.SetValue("founder_1_precent_fond", founder_1_precent_fond);


                    file.SaveDocument(name_firma_ru + "_List_1.docx");
                   
                    file.Dispose();
                   
                    file = new Docx("template_list_a.docx");
                    file.SetValue("founder_2_type_document", founder_2_type_document);
                    file.SetValue("founder_2_citizen", founder_2_citizen);
                    file.SetValue("founder_2_country", founder_2_country);

                    file.SetValue("founder_2_firstname", founder_2_firstname);
                    file.SetValue("founder_2_name", founder_2_name);
                    file.SetValue("founder_2_lastname", founder_2_lastname);
                    file.SetValue("founder_2_sex", founder_2_sex);
                    file.SetValue("founder_2_datebirthd", founder_2_datebirthd);
                    file.SetValue("founder_2_placebirthd", founder_2_placebirthd);
                    file.SetValue("founder_2_index", founder_2_index);
                    file.SetValue("founder_2_region", founder_2_region);
                    file.SetValue("founder_2_district", this.founder_2_district);
                    switch (locality_founder_2)
                    {
                        case "city_founder_2":
                            file.SetValue("founder_2_city", founder_2_name_city);
                            file.SetValue("founder_2_minicity", "");
                            file.SetValue("founder_2_village", "");
                            break;
                        case "minicity_founder_2":
                            file.SetValue("founder_2_minicity", founder_2_name_city);
                            file.SetValue("founder_2_village", "");
                            file.SetValue("founder_2_city", "");

                            break;
                        case "village_founder_2":
                            file.SetValue("founder_2_village", founder_2_name_city);
                            file.SetValue("founder_2_city", "");
                            file.SetValue("founder_2_minicity", "");
                            break;
                    }
                    file.SetValue("founder_2_street", founder_2_street);
                    file.SetValue("founder_2_korpus", founder_2_korpus);

                    file.SetValue("founder_2_number_house", founder_2_number_house);
                    file.SetValue("founder_2_app", founder_2_app);
                    file.SetValue("founder_2_id_passport", founder_2_id_passport);
                    file.SetValue("founder_2_number_passport", founder_2_number_passport);
                    file.SetValue("founder_2_date_passport", founder_2_date_passport);
                    file.SetValue("founder_2_valid_passport", founder_2_valid_passport);
                    file.SetValue("founder_2_registration_passport", founder_2_registration_passport);

                    file.SetValue("founder_2_code", founder_2_code);
                    file.SetValue("founder_2_phone", founder_2_phone);

                    file.SetValue("founder_2_fond", founder_2_fond);
                    file.SetValue("founder_2_precent_fond", founder_2_precent_fond);


                    file.SaveDocument(name_firma_ru + "_List_2.docx");
                   
                    file.Dispose();
                   
                    file = new Docx("template_list_c.docx");
                    file.SetValue("founder_3_type_document", founder_3_type_document);
                 
                    file.SetValue("founder_3_citizen", founder_3_citizen);
                   
                    file.SetValue("founder_3_country", founder_3_country);
                    file.SetValue("founder_3_firstname", founder_3_firstname);
                    file.SetValue("founder_3_name", founder_3_name);
                    file.SetValue("founder_3_lastname", founder_3_lastname);
                    file.SetValue("founder_3_sex", founder_3_sex);
                    file.SetValue("founder_3_datebirthd", founder_3_datebirthd);
                    file.SetValue("founder_3_placebirthd", founder_3_placebirthd);
                    file.SetValue("founder_3_index", founder_3_index);
                    file.SetValue("founder_3_region", founder_3_region);
                    file.SetValue("founder_3_district", founder_3_district);
                    switch (locality_founder_3)
                    {
                        case "city_founder_3":
                            file.SetValue("founder_3_city", founder_3_name_city);
                            file.SetValue("founder_3_minicity", "");
                            file.SetValue("founder_3_village", "");
                            break;
                        case "minicity_founder_3":
                            file.SetValue("founder_3_minicity", founder_3_name_city);
                            file.SetValue("founder_3_village", "");
                            file.SetValue("founder_3_city", "");

                            break;
                        case "village_founder_3":
                            file.SetValue("founder_3_village", founder_3_name_city);
                            file.SetValue("founder_3_city", "");
                            file.SetValue("founder_3_minicity", "");
                            break;
                    }
                    file.SetValue("founder_3_street", founder_3_street);
                    file.SetValue("founder_3_korpus", founder_3_korpus);

                    file.SetValue("founder_3_number_house", founder_3_number_house);
                    file.SetValue("founder_3_app", founder_3_app);
                    file.SetValue("founder_3_id_passport", founder_3_id_passport);
                    file.SetValue("founder_3_number_passport", founder_3_number_passport);
                    file.SetValue("founder_3_date_passport", founder_3_date_passport);
                    file.SetValue("founder_3_valid_passport", founder_3_valid_passport);
                    file.SetValue("founder_3_registration_passport", founder_3_registration_passport);

                    file.SetValue("founder_3_code", founder_3_code);
                    file.SetValue("founder_3_phone", founder_3_phone);

                    file.SetValue("founder_3_fond", founder_3_fond);
                    file.SetValue("founder_3_precent_fond", founder_3_precent_fond);
                    file.SaveDocument(name_firma_ru + "_List3.docx"); //
                    file.Dispose();
                   
                    file = new Docx("template_protocol_1_founder_3.docx");
                    file.SetValue("name_firma_ru", name_firma_ru);

                    file.SetValue("oked_name", oked_name);
                    file.SetValue("oked", oked);

                    file.SetValue("all_fond", all_fond);


                    file.SetValue("date_protocol", date_protocol);
                    file.SetValue("date_protocol_2", date_protocol_2);


                    if (region_ur.ToLower() == "нет")
                    {
                        file.SetValue("region_ur", " ");

                    }
                    else
                    {
                        file.SetValue("region_ur", " " + region_ur + " область,");
                    }

                    if (district_ur.ToLower() == "нет")
                    {

                        file.SetValue("district_ur", " ");
                    }
                    else
                    {
                        file.SetValue("district_ur", " " + district_ur + " район,");
                    }
                    switch (locality_ur)
                    {
                        case "city_ur":
                            file.SetValue("city_ur", "город " + name_city_ur + ",");

                            break;
                        case "minicity_ur":
                            file.SetValue("city_ur", "посёлок " + name_city_ur + ",");


                            break;
                        case "village_ur":
                            file.SetValue("city_ur", "сельский совет " + name_city_ur + ",");

                            break;
                    }
                    file.SetValue("street_ur", street_ur);
                    file.SetValue("house_ur", ", д. " + house_ur);

                    if (korpus_ur.ToLower() == "нет")
                    {

                        file.SetValue("korpus_ur", ", ");
                    }
                    else
                    {
                        file.SetValue("korpus_ur", ", корпус: " + korpus_ur + ",");
                    }
                    file.SetValue("office_ur", office_ur);
                    file.SetValue("founder_1_type_document", founder_1_type_document);
                    file.SetValue("founder_2_type_document", founder_2_type_document);
                    file.SetValue("founder_3_type_document", founder_3_type_document);
                    file.SetValue("founder_1_citizen", founder_1_citizen);
                    file.SetValue("founder_2_citizen", founder_2_citizen);
                    file.SetValue("founder_3_citizen", founder_3_citizen);
                    file.SetValue("founder_1_country", founder_1_country);
                    file.SetValue("founder_2_country", founder_2_country);
                    file.SetValue("founder_3_country", founder_3_country);
                    file.SetValue("founder_1_firstname", founder_1_firstname);
                    file.SetValue("founder_1_name", founder_1_name);
                    file.SetValue("founder_1_lastname", founder_1_lastname);

                    file.SetValue("founder_1_datebirdth", founder_1_datebirthd);
                    file.SetValue("founder_1_index", founder_1_index);



                    if (founder_1_region.ToLower() == "нет")
                    {
                        file.SetValue("founder_1_region", " ");

                    }
                    else
                    {
                        file.SetValue("founder_1_region", founder_1_region + " область,");
                    }

                    if (this.founder_1_district.ToLower() == "нет")
                    {

                        file.SetValue("founder_1_district", " ");
                    }
                    else
                    {
                        file.SetValue("founder_1_district", this.founder_1_district + " район,");
                    }
                    file.SetValue("founder_1_street", founder_1_street + ", ");
                    switch (locality_founder_1)
                    {
                        case "city_founder_1":
                            file.SetValue("founder_1_city", "город " + founder_1_name_city + ",");

                            break;
                        case "minicity_founder_1":
                            file.SetValue("founder_1_city", "посёлок " + founder_1_name_city + ",");


                            break;
                        case "village_founder_1":
                            file.SetValue("founder_1_city", "сельский совет " + founder_1_name_city + ",");

                            break;
                    }


                    if (founder_1_korpus.ToLower() == "нет")
                    {
                        file.SetValue("founder_1_korpus", "");

                    }
                    else
                    {
                        file.SetValue("founder_1_korpus", " корпус: " + founder_1_korpus + ",");
                    }
                    file.SetValue("founder_1_house", " д. " + founder_1_number_house);

                    if (founder_1_app.ToLower() == "")
                    {
                        file.SetValue("founder_1_app", "");
                    }
                    else
                    {
                        file.SetValue("founder_1_app", ", кв. " + founder_1_app);

                    }
                    file.SetValue("founder_1_idpassport", founder_1_id_passport);
                    file.SetValue("founder_1_numberpassport", founder_1_number_passport);
                    file.SetValue("founder_1_datepassport", founder_1_date_passport);
                    file.SetValue("founder_1_registrationpassport", founder_1_registration_passport);

                    file.SetValue("founder_1_fond", founder_1_fond);
                    file.SetValue("founder_1_precent_fond", founder_1_precent_fond);

                    file.SetValue("founder_2_firstname", founder_2_firstname);
                    file.SetValue("founder_2_name", founder_2_name);
                    file.SetValue("founder_2_lastname", founder_2_lastname);

                    file.SetValue("founder_2_datebirdth", founder_2_datebirthd);
                    file.SetValue("founder_2_index", founder_2_index);


                    file.SetValue("founder_2_idpassport", founder_2_id_passport);
                    file.SetValue("founder_2_numberpassport", founder_2_number_passport);
                    file.SetValue("founder_2_datepassport", founder_2_date_passport);
                    file.SetValue("founder_2_registrationpassport", founder_2_registration_passport);
                    file.SetValue("founder_2_fond", founder_2_fond);
                    file.SetValue("founder_2_precent_fond", founder_2_precent_fond);


                    if (founder_2_region.ToLower() == "нет")
                    {
                        file.SetValue("founder_2_region", "");

                    }
                    else
                    {
                        file.SetValue("founder_2_region", founder_2_region + " область,");
                    }

                    if (founder_2_district.ToLower() == "нет")
                    {

                        file.SetValue("founder_2_district", "");
                    }
                    else
                    {
                        file.SetValue("founder_2_district", founder_2_district + " район,");
                    }
                    file.SetValue("founder_2_street", founder_2_street + ", ");
                    switch (locality_founder_2)
                    {
                        case "city_founder_2":
                            file.SetValue("founder_2_city", "город " + founder_2_name_city + ",");

                            break;
                        case "minicity_founder_2":
                            file.SetValue("founder_2_city", "посёлок " + founder_2_name_city + ",");


                            break;
                        case "village_founder_2":
                            file.SetValue("founder_2_city", "сельский совет " + founder_2_name_city + ",");

                            break;
                    }


                    if (founder_2_korpus.ToLower() == "нет")
                    {
                        file.SetValue("founder_2_korpus", "");

                    }
                    else
                    {
                        file.SetValue("founder_2_korpus", " корпус: " + founder_2_korpus + ",");
                    }
                    file.SetValue("founder_2_house", " д. " + founder_2_number_house);

                    if (founder_2_app == "частный дом")
                    {
                        file.SetValue("founder_2_app", "");
                    }
                    else
                    {
                        file.SetValue("founder_2_app", ", кв." + founder_2_app);
                    }

                    file.SetValue("founder_3_firstname", founder_3_firstname);
                    file.SetValue("founder_3_name", founder_3_name);
                    file.SetValue("founder_3_lastname", founder_3_lastname);

                    file.SetValue("founder_3_datebirdth", founder_3_datebirthd);
                    file.SetValue("founder_3_index", founder_3_index);


                    file.SetValue("founder_3_idpassport", founder_3_id_passport);
                    file.SetValue("founder_3_numberpassport", founder_3_number_passport);
                    file.SetValue("founder_3_datepassport", founder_3_date_passport);
                    file.SetValue("founder_3_registrationpassport", founder_3_registration_passport);
                    file.SetValue("founder_3_fond", founder_3_fond);
                    file.SetValue("founder_3_precent_fond", founder_3_precent_fond);


                    if (founder_3_region.ToLower() == "нет")
                    {
                        file.SetValue("founder_3_region", "");

                    }
                    else
                    {
                        file.SetValue("founder_3_region", founder_3_region + " область,");
                    }

                    if (founder_3_district.ToLower() == "нет")
                    {

                        file.SetValue("founder_3_district", "");
                    }
                    else
                    {
                        file.SetValue("founder_3_district", founder_3_district + " район,");
                    }
                    file.SetValue("founder_3_street", founder_3_street + ", ");
                    switch (locality_founder_3)
                    {
                        case "city_founder_3":
                            file.SetValue("founder_3_city", "город " + founder_3_name_city + ",");

                            break;
                        case "minicity_founder_3":
                            file.SetValue("founder_3_city", "посёлок " + founder_3_name_city + ",");


                            break;
                        case "village_founder_3":
                            file.SetValue("founder_3_city", "сельский совет " + founder_3_name_city + ",");

                            break;
                    }


                    if (founder_3_korpus.ToLower() == "нет")
                    {
                        file.SetValue("founder_3_korpus", "");

                    }
                    else
                    {
                        file.SetValue("founder_3_korpus", " корпус: " + founder_3_korpus + ",");
                    }
                    file.SetValue("founder_3_house", " д. " + founder_3_number_house);

                    if (founder_3_app == "частный дом")
                    {
                        file.SetValue("founder_3_app", "");
                    }
                    else
                    {
                        file.SetValue("founder_3_app", ", кв." + founder_3_app);
                    }

                    file.SaveDocument(name_firma_ru + "_Протокол№1.docx");
                    file.Dispose();
                   

                    file = new Docx("template_protocol_2_founder_3.docx");
                    file.SetValue("founder_1_type_document", founder_1_type_document);
                    file.SetValue("founder_2_type_document", founder_2_type_document);
                    file.SetValue("founder_3_type_document", founder_3_type_document);
                    file.SetValue("founder_1_citizen", founder_1_citizen);
                    file.SetValue("founder_2_citizen", founder_2_citizen);
                    file.SetValue("founder_3_citizen", founder_3_citizen);
                    file.SetValue("founder_1_country", founder_1_country + ", ");
                    file.SetValue("founder_2_country", founder_2_country + ", ");
                    file.SetValue("founder_3_country", founder_3_country + ", ");
                    file.SetValue("director_type_document", director_type_document);

                    file.SetValue("firstname_director", firstname_director);
                    file.SetValue("name_director", name_director);
                    file.SetValue("lastname_director", lastname_director);
                    file.SetValue("datebirdth_director", datebirthd_director);
                    file.SetValue("placebirthd_director", placebirthd_director);
                    file.SetValue("index_director", index_director);
                    file.SetValue("director_country", director_country);
                    if (region_director.ToLower() == "нет")
                    {
                        file.SetValue("region_director", "");

                    }
                    else
                    {
                        file.SetValue("region_director", region_director + " область,");
                    }

                    if (district_director.ToLower() == "нет")
                    {

                        file.SetValue("district_director", "");
                    }
                    else
                    {
                        file.SetValue("district_director", district_director + " район,");
                    }
                    file.SetValue("street_director", ", " + street_director);


                    switch (locality_director)
                    {
                        case "city_director":
                            file.SetValue("city_director", "город " + name_city_director);

                            break;
                        case "minicity_director":
                            file.SetValue("city_director", "посёлок " + name_city_director);


                            break;
                        case "village_director":
                            file.SetValue("city_director", "сельский совет " + name_city_director);

                            break;
                    }


                    if (korpus_director.ToLower() == "нет")
                    {
                        file.SetValue("korpus_director", "");

                    }
                    else
                    {
                        file.SetValue("korpus_director", ", корпус: " + korpus_director + ",");
                    }
                    file.SetValue("house_director", ", д. " + house_director);

                    if (app_director.ToLower() == "частный дом")
                    {
                        file.SetValue("app_director", "");
                    }
                    else
                    {
                        file.SetValue("app_director", ", кв. " + app_director);

                    }
                    file.SetValue("idpassport_director", idpassport_director);
                    file.SetValue("numberpassport_director", numberpassport_director);
                    file.SetValue("datepassport_director", datapassport_director);
                    file.SetValue("validpassport_director", validpassport_director);
                    file.SetValue("registrationpassport_director", authority_passport_director);


                    file.SetValue("name_firma_ru", name_firma_ru);

                    file.SetValue("oked_name", oked_name);
                    file.SetValue("oked", oked);


                    file.SetValue("date_protocol", date_protocol);
                    file.SetValue("date_protocol_2", date_protocol_2);




                    if (region_ur.ToLower() == "нет")
                    {
                        file.SetValue("region_ur", ", ");

                    }
                    else
                    {
                        file.SetValue("region_ur", region_ur + " область,");
                    }

                    if (district_ur.ToLower() == "нет")
                    {

                        file.SetValue("district_ur", " ");
                    }
                    else
                    {
                        file.SetValue("district_ur", district_ur + " район,");
                    }
                    switch (locality_ur)
                    {
                        case "city_ur":
                            file.SetValue("city_ur", "город " + name_city_ur);

                            break;
                        case "minicity_ur":
                            file.SetValue("city_ur", "посёлок " + name_city_ur);


                            break;
                        case "village_ur":
                            file.SetValue("city_ur", "сельский совет " + name_city_ur);

                            break;
                    }
                    file.SetValue("office_ur", office_ur);
                    file.SetValue("founder_1_firstname", founder_1_firstname);
                    file.SetValue("founder_1_name", founder_1_name);
                    file.SetValue("founder_1_lastname", founder_1_lastname);

                    file.SetValue("founder_1_datebirthd", founder_1_datebirthd);
                    file.SetValue("founder_1_index", founder_1_index);

                    if (founder_1_region.ToLower() == "нет")
                    {
                        file.SetValue("founder_1_region", "");

                    }
                    else
                    {
                        file.SetValue("founder_1_region", founder_1_region + " область, ");
                    }

                    if (this.founder_1_district.ToLower() == "нет")
                    {

                        file.SetValue("founder_1_district", "");
                    }
                    else
                    {
                        file.SetValue("founder_1_district", this.founder_1_district + " район,");
                    }
                    file.SetValue("founder_1_street", founder_1_street + ", ");
                    switch (locality_founder_1)
                    {
                        case "city_founder_1":
                            file.SetValue("founder_1_city", " город " + founder_1_name_city + ", ");

                            break;
                        case "minicity_founder_1":
                            file.SetValue("founder_1_city", " посёлок " + founder_1_name_city + ",");


                            break;
                        case "village_founder_1":
                            file.SetValue("founder_1_city", " сельский совет " + founder_1_name_city + ",");

                            break;
                    }


                    if (founder_1_korpus.ToLower() == "нет")
                    {
                        file.SetValue("founder_1_korpus", " ");

                    }
                    else
                    {
                        file.SetValue("founder_1_korpus", " корпус: " + founder_1_korpus + ", ");
                    }
                    file.SetValue("founder_1_house", " д. " + founder_1_number_house);
                    if (founder_1_app == "частный дом")
                    {
                        file.SetValue("founder_1_app", "");
                    }
                    else
                    {
                        file.SetValue("founder_1_app", ", кв. " + founder_1_app);

                    }


                    file.SetValue("founder_1_idpassport", founder_1_id_passport);
                    file.SetValue("founder_1_numberpassport", founder_1_number_passport);
                    file.SetValue("founder_1_datepassport", founder_1_date_passport);
                    file.SetValue("founder_1_registrationpassport", founder_1_registration_passport);


                    file.SetValue("founder_2_firstname", founder_2_firstname);
                    file.SetValue("founder_2_name", founder_2_name);
                    file.SetValue("founder_2_lastname", founder_2_lastname);

                    file.SetValue("founder_2_datebirthd", founder_2_datebirthd);
                    file.SetValue("founder_2_index", founder_2_index);




                    if (founder_2_region.ToLower() == "нет")
                    {
                        file.SetValue("founder_2_region", "");

                    }
                    else
                    {
                        file.SetValue("founder_2_region", founder_2_region + " область, ");
                    }

                    if (founder_2_district.ToLower() == "нет")
                    {

                        file.SetValue("founder_2_district", "");
                    }
                    else
                    {
                        file.SetValue("founder_2_district", founder_2_district + " район,");
                    }
                    file.SetValue("founder_2_street", founder_2_street + ", ");
                    switch (locality_founder_2)
                    {
                        case "city_founder_2":
                            file.SetValue("founder_2_city", " город " + founder_2_name_city + ", ");

                            break;
                        case "minicity_founder_2":
                            file.SetValue("founder_2_city", " посёлок " + founder_2_name_city + ",");


                            break;
                        case "village_founder_2":
                            file.SetValue("founder_2_city", " сельский совет " + founder_2_name_city + ",");

                            break;
                    }


                    if (founder_2_korpus.ToLower() == "нет")
                    {
                        file.SetValue("founder_2_korpus", " ");

                    }
                    else
                    {
                        file.SetValue("founder_2_korpus", " корпус: " + founder_2_korpus + ", ");
                    }
                    file.SetValue("founder_2_house", " д. " + founder_2_number_house);
                    if (founder_2_app == "частный дом")
                    {
                        file.SetValue("founder_2_app", "");
                    }
                    else
                    {
                        file.SetValue("founder_2_app", ", кв. " + founder_2_app);

                    }
                    file.SetValue("founder_2_idpassport", founder_2_id_passport);
                    file.SetValue("founder_2_numberpassport", founder_2_number_passport);
                    file.SetValue("founder_2_datepassport", founder_2_date_passport);
                    file.SetValue("founder_2_registrationpassport", founder_2_registration_passport);



                    file.SetValue("founder_3_firstname", founder_3_firstname);
                    file.SetValue("founder_3_name", founder_3_name);
                    file.SetValue("founder_3_lastname", founder_3_lastname);

                    file.SetValue("founder_3_datebirthd", founder_3_datebirthd);
                    file.SetValue("founder_3_index", founder_3_index);




                    if (founder_3_region.ToLower() == "нет")
                    {
                        file.SetValue("founder_3_region", "");

                    }
                    else
                    {
                        file.SetValue("founder_3_region", founder_3_region + " область, ");
                    }

                    if (founder_3_district.ToLower() == "нет")
                    {

                        file.SetValue("founder_3_district", "");
                    }
                    else
                    {
                        file.SetValue("founder_3_district", founder_3_district + " район,");
                    }
                    file.SetValue("founder_3_street", founder_3_street + ", ");
                    switch (locality_founder_3)
                    {
                        case "city_founder_3":
                            file.SetValue("founder_3_city", " город " + founder_3_name_city + ", ");

                            break;
                        case "minicity_founder_3":
                            file.SetValue("founder_3_city", " посёлок " + founder_3_name_city + ",");


                            break;
                        case "village_founder_3":
                            file.SetValue("founder_3_city", " сельский совет " + founder_3_name_city + ",");

                            break;
                    }


                    if (founder_3_korpus.ToLower() == "нет")
                    {
                        file.SetValue("founder_3_korpus", " ");

                    }
                    else
                    {
                        file.SetValue("founder_3_korpus", " корпус: " + founder_3_korpus + ", ");
                    }
                    file.SetValue("founder_3_house", " д. " + founder_3_number_house);
                    if (founder_3_app == "частный дом")
                    {
                        file.SetValue("founder_3_app", "");
                    }
                    else
                    {
                        file.SetValue("founder_3_app", ", кв. " + founder_3_app);

                    }
                    file.SetValue("founder_3_idpassport", founder_3_id_passport);
                    file.SetValue("founder_3_numberpassport", founder_3_number_passport);
                    file.SetValue("founder_3_datepassport", founder_3_date_passport);
                    file.SetValue("founder_3_registrationpassport", founder_3_registration_passport);


                    file.SaveDocument(name_firma_ru + "_Протокол№2.docx");
                    file.Dispose();
                   

                    file = new Docx("templateustav_ooo_founder_3.docx");
                    file.SetValue("name_firma_ru", name_firma_ru);
                    file.SetValue("date_protocol", date_protocol);
                    file.SetValue("year", DateTime.Now.Year.ToString());
                    file.SetValue("name_firma_by", name_firma_by);
                    file.SetValue("founder_1_firstname", founder_1_firstname);
                    file.SetValue("founder_1_name", founder_1_name);
                    file.SetValue("founder_1_lastname", founder_1_lastname);
                    file.SetValue("founder_1_datebirthd", founder_1_datebirthd);
                    file.SetValue("founder_1_placebirthd", founder_1_placebirthd);
                    file.SetValue("founder_1_index", founder_1_index);
                    file.SetValue("founder_1_type_document", founder_1_type_document);
                    file.SetValue("founder_2_type_document", founder_2_type_document);
                    file.SetValue("founder_1_citizen", founder_1_citizen);
                    file.SetValue("founder_2_citizen", founder_2_citizen);
                    file.SetValue("founder_1_country", founder_1_country);
                    file.SetValue("founder_2_country", founder_2_country);
                    file.SetValue("founder_3_citizen", founder_3_citizen);
                    file.SetValue("founder_3_country", founder_3_country);
                    file.SetValue("founder_3_type_document", founder_3_type_document);
                    if (founder_1_region.ToLower() == "нет")
                    {
                        file.SetValue("founder_1_region", "");

                    }
                    else
                    {
                        file.SetValue("founder_1_region", founder_1_region + " область,");
                    }

                    if (founder_1_district.ToLower() == "нет")
                    {

                        file.SetValue("founder_1_district", "");
                    }
                    else
                    {
                        file.SetValue("founder_1_district", founder_1_district + " район,");
                    }
                    file.SetValue("founder_1_street", founder_1_street + ", ");
                    switch (locality_founder_1)
                    {
                        case "city_founder_1":
                            file.SetValue("founder_1_city", "город " + founder_1_name_city + ", ");

                            break;
                        case "minicity_founder_1":
                            file.SetValue("founder_1_city", "посёлок " + founder_1_name_city + ", ");


                            break;
                        case "village_founder_1":
                            file.SetValue("founder_1_city", "сельский совет " + founder_1_name_city + ", ");

                            break;
                    }


                    if (founder_1_korpus.ToLower() == "нет")
                    {
                        file.SetValue("founder_1_korpus", "");

                    }
                    else
                    {
                        file.SetValue("founder_1_korpus", "корпус: " + founder_1_korpus + ", ");
                    }
                    file.SetValue("founder_1_house", " д. " + founder_1_number_house);

                    if (founder_1_app == "частный дом")
                    {
                        file.SetValue("founder_1_app", "");
                    }
                    else
                    {
                        file.SetValue("founder_1_app", ", кв. " + founder_1_app);

                    }
                    file.SetValue("founder_1_id_passport", founder_1_id_passport);
                    file.SetValue("founder_1_number_passport", founder_1_number_passport);
                    file.SetValue("founder_1_date_passport", founder_1_date_passport);
                    file.SetValue("founder_1_valid_passport", founder_1_valid_passport);
                    file.SetValue("founder_1_registration_passport", founder_1_registration_passport);
                    file.SetValue("founder_1_fond", founder_1_fond);
                    file.SetValue("founder_1_precent_fond", founder_1_precent_fond);

                    file.SetValue("founder_2_firstname", founder_2_firstname);
                    file.SetValue("founder_2_name", founder_2_name);
                    file.SetValue("founder_2_lastname", founder_2_lastname);
                    file.SetValue("founder_2_datebirthd", founder_2_datebirthd);
                    file.SetValue("founder_2_placebirthd", founder_2_placebirthd);
                    file.SetValue("founder_2_index", founder_2_index);



                    if (founder_2_region.ToLower() == "нет")
                    {
                        file.SetValue("founder_2_region", "");

                    }
                    else
                    {
                        file.SetValue("founder_2_region", founder_2_region + " область, ");
                    }

                    if (founder_2_district.ToLower() == "нет")
                    {

                        file.SetValue("founder_2_district", "");
                    }
                    else
                    {
                        file.SetValue("founder_2_district", founder_2_district + " район,");
                    }
                    file.SetValue("founder_2_street", founder_2_street + ", ");
                    switch (locality_founder_2)
                    {
                        case "city_founder_2":
                            file.SetValue("founder_2_city", " город " + founder_2_name_city + ", ");

                            break;
                        case "minicity_founder_2":
                            file.SetValue("founder_2_city", " посёлок " + founder_2_name_city + ",");


                            break;
                        case "village_founder_2":
                            file.SetValue("founder_2_city", " сельский совет " + founder_2_name_city + ",");

                            break;
                    }


                    if (founder_2_korpus.ToLower() == "нет")
                    {
                        file.SetValue("founder_2_korpus", " ");

                    }
                    else
                    {
                        file.SetValue("founder_2_korpus", " корпус: " + founder_2_korpus + ", ");
                    }
                    file.SetValue("founder_2_house", " д. " + founder_2_number_house);
                    if (founder_2_app == "частный дом")
                    {
                        file.SetValue("founder_2_app", "");
                    }
                    else
                    {
                        file.SetValue("founder_2_app", ", кв. " + founder_2_app);

                    }
                    file.SetValue("founder_2_id_passport", founder_2_id_passport);
                    file.SetValue("founder_2_number_passport", founder_2_number_passport);
                    file.SetValue("founder_2_date_passport", founder_2_date_passport);
                    file.SetValue("founder_2_valid_passport", founder_2_valid_passport);
                    file.SetValue("founder_2_registration_passport", founder_2_registration_passport);
                    file.SetValue("founder_2_fond", founder_2_fond);
                    file.SetValue("founder_2_precent_fond", founder_2_precent_fond);

                    file.SetValue("founder_3_firstname", founder_3_firstname);
                    file.SetValue("founder_3_name", founder_3_name);
                    file.SetValue("founder_3_lastname", founder_3_lastname);
                    file.SetValue("founder_3_datebirthd", founder_3_datebirthd);
                    file.SetValue("founder_3_placebirthd", founder_3_placebirthd);
                    file.SetValue("founder_3_index", founder_3_index);



                    if (founder_3_region.ToLower() == "нет")
                    {
                        file.SetValue("founder_3_region", "");

                    }
                    else
                    {
                        file.SetValue("founder_3_region", founder_3_region + " область, ");
                    }

                    if (founder_3_district.ToLower() == "нет")
                    {

                        file.SetValue("founder_3_district", "");
                    }
                    else
                    {
                        file.SetValue("founder_3_district", founder_3_district + " район,");
                    }
                    file.SetValue("founder_3_street", founder_3_street + ", ");
                    switch (locality_founder_3)
                    {
                        case "city_founder_3":
                            file.SetValue("founder_3_city", " город " + founder_3_name_city + ", ");

                            break;
                        case "minicity_founder_3":
                            file.SetValue("founder_3_city", " посёлок " + founder_3_name_city + ",");


                            break;
                        case "village_founder_3":
                            file.SetValue("founder_3_city", " сельский совет " + founder_3_name_city + ",");

                            break;
                    }


                    if (founder_3_korpus.ToLower() == "нет")
                    {
                        file.SetValue("founder_3_korpus", " ");

                    }
                    else
                    {
                        file.SetValue("founder_3_korpus", " корпус: " + founder_3_korpus + ", ");
                    }
                    file.SetValue("founder_3_house", " д. " + founder_3_number_house + ",");
                    if (founder_3_app == "частный дом")
                    {
                        file.SetValue("founder_3_app", "");
                    }
                    else
                    {
                        file.SetValue("founder_3_app", ", кв. " + founder_3_app);

                    }
                    file.SetValue("founder_3_id_passport", founder_3_id_passport);
                    file.SetValue("founder_3_number_passport", founder_3_number_passport);
                    file.SetValue("founder_3_date_passport", founder_3_date_passport);
                    file.SetValue("founder_3_valid_passport", founder_3_valid_passport);
                    file.SetValue("founder_3_registration_passport", founder_3_registration_passport);
                    file.SetValue("founder_3_fond", founder_3_fond);
                    file.SetValue("founder_3_precent_fond", founder_3_precent_fond);

                    file.SetValue("all_fond", all_fond);


                    if (region_ur.ToLower() == "нет")
                    {
                        file.SetValue("region_ur", "");

                    }
                    else
                    {
                        file.SetValue("region_ur", region_ur + " область,");
                    }

                    if (district_ur.ToLower() == "нет")
                    {

                        file.SetValue("district_ur", "");
                    }
                    else
                    {
                        file.SetValue("district_ur", district_ur + " район,");
                    }
                    switch (locality_ur)
                    {
                        case "city_ur":
                            file.SetValue("city_ur", " город " + name_city_ur + ",");

                            break;
                        case "minicity_ur":
                            file.SetValue("city_ur", " посёлок " + name_city_ur + ",");


                            break;
                        case "village_ur":
                            file.SetValue("city_ur", " сельский совет " + name_city_ur + ",");

                            break;
                    }
                    file.SetValue("street_ur", street_ur + ", ");
                    file.SetValue("house_ur", " д. " + house_ur + ", ");

                    if (korpus_ur.ToLower() == "нет")
                    {

                        file.SetValue("korpus_ur", " ");
                    }
                    else
                    {
                        file.SetValue("korpus_ur", " корпус: " + korpus_ur + ", ");
                    }
                    file.SetValue("office_ur", office_ur);
                    file.SaveDocument(name_firma_ru + "_Устав.docx");//
                    file.Dispose();
                  
                    file = new Docx("template_gos.docx");
                    file.SetValue("registration", registering_authority);
                    file.SetValue("soglas", soglas);
                    file.SetValue("name_firma_ru", name_firma_ru);
                    file.SetValue("name_firma_by", name_firma_by);
                    file.SetValue("date_protocol_2", date_protocol_2);

                    file.SetValue("firstname_director", firstname_director);
                    file.SetValue("name_director", name_director);
                    file.SetValue("lastname_director", lastname_director);
                    file.SetValue("datebirthd_director", datebirthd_director);
                    file.SetValue("placebirthd_director", placebirthd_director);
                    file.SetValue("director_country", director_country);
                    file.SetValue("index_director", index_director);
                    file.SetValue("region_director", region_director);
                    file.SetValue("district_director", district_director);
                    switch (locality_director)
                    {
                        case "city_director":
                            file.SetValue("city_director", name_city_director);
                            file.SetValue("minicity_director", "");
                            file.SetValue("village_director", "");
                            break;
                        case "minicity_director":
                            file.SetValue("minicity_director", name_city_director);
                            file.SetValue("city_director", "");
                            file.SetValue("village_director", "");
                            break;
                        case "village_director":
                            file.SetValue("village_director", name_city_director);
                            file.SetValue("minicity_director", "");
                            file.SetValue("city_director", "");
                            break;
                    }

                    file.SetValue("street_director", street_director);

                    file.SetValue("korpus_director", korpus_director);

                    file.SetValue("house_director", house_director);
                    file.SetValue("app_director", app_director);
                    file.SetValue("idpassport_director", idpassport_director);
                    file.SetValue("numberpassport_director", numberpassport_director);
                    file.SetValue("datepassport_director", datapassport_director);
                    file.SetValue("validpassport_director", validpassport_director);
                    file.SetValue("registrationpassport_director", authority_passport_director);
                    file.SetValue("code_director", code_director);
                    file.SetValue("phone_director", phone_director);
                    file.SetValue("oked_name", oked_name);
                    file.SetValue("oked", oked);
                    file.SetValue("all_fond", all_fond);
                    file.SetValue("region_ur", region_ur);
                    file.SetValue("district_ur", district_ur);
                    switch (locality_ur)
                    {
                        case "city_ur":
                            file.SetValue("city_ur", name_city_ur);
                            file.SetValue("minicity_ur", "");
                            file.SetValue("village_ur", "");
                            break;
                        case "minicity_ur":
                            file.SetValue("minicity_ur", name_city_ur);
                            file.SetValue("city_ur", "");
                            file.SetValue("village_ur", "");
                            break;
                        case "village_ur":
                            file.SetValue("village_ur", name_city_ur);
                            file.SetValue("minicity_ur", "");
                            file.SetValue("city_ur", "");
                            break;
                    }
                    file.SetValue("street_ur", street_ur);
                    file.SetValue("index_ur", index_ur);
                    file.SetValue("house_ur", house_ur);
                    file.SetValue("korpus_ur", korpus_ur);
                    file.SetValue("office_ur", office_ur);


                    file.SetValue("director_type_document", director_type_document);
                    file.SaveDocument(name_firma_ru + "_Регистрация.docx");
                    file.Dispose();
                   

                    MessageBox.Show("Завершено!");
                    break;
                    
                    
            }

          }
        }
    }

