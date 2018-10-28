import word_base
import time
import config

def get_handler_data_to_cutpane_Shear_plate(file_path):

    '''将数据渲染后放到剪切板上,并存到一个临时文件中'''

    word_app = word_base.WordWrap(file_path)

    word_app_2=word_base.WordWrap()
    word_app_3=word_base.WordWrap()

    word_app.show()
    word_app_2.show()
    word_app_3.show()


    for data_tuple in config.data_list:

        word_app.all_copy()
        time.sleep(1)
        word_app_2.paste_end()

        for i in range(len(data_tuple)):
            # print(i,end='   ')
            time.sleep(1)
            word_app_2.replace_word(config.context_mark[i],data_tuple[i])
        # print()
        word_app_2.all_cut()
        time.sleep(1)
        word_app_3.paste_end()

    word_app_4=word_base.WordWrap(config.dest_file)
    # word_app_4.show()
    word_app_3.all_copy()
    time.sleep(1)
    word_app_4.word_before(config.word_before)
    time.sleep(1)
    word_app_4.paste_origin()
    time.sleep(1)
    word_app_4.update_catalog()


    word_app_4.save()
    time.sleep(1)
    word_app_3.saveAs(r'%s\result_%s.docx' %( config.temp_dir,time.time()))
    time.sleep(1)
    word_app_4.close()
    time.sleep(1)
    word_app_3.close()
    time.sleep(1)
    word_app.close()
    time.sleep(1)
    word_app_2.close()


def main():

    get_handler_data_to_cutpane_Shear_plate(config.template_file)
    print('ok')


if __name__ == '__main__':
    main()