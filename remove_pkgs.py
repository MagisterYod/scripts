import subprocess
import time

if __name__ == '__main__':

    while True:

        run_app = str(
            subprocess.Popen(['adb', 'devices'],
                             stdout=subprocess.PIPE).communicate()[0]).split('n')[1].split('r')[0]
        packages_6200 = [
            'com.google.android.youtube',
            'com.google.android.apps.googleassistant',
            'com.google.android.apps.youtube.music',
            'com.google.android.keep',
            'com.android.musicfx',
            'com.google.android.apps.maps',
            'com.google.android.apps.photos',
            'com.google.android.overlay.gmsconfig.photos',
            'com.sprd.providers.photos',
            'com.blackview.gamemode',
            'com.google.android.apps.translate',
            'com.google.android.videos',
            'com.android.fmradio',
            'com.google.android.apps.tachyon',
            'ru.yandex.yandexmaps',
            'com.yandex.searchapp',
            'com.yandex.browser',
            'com.google.android.overlay.modules.healthfitness.forframework',
            'club.dexp.minimarket2',
            'com.blackview.easytrans',
            'com.google.android.apps.fitness',
            'cn.wps.moffice_eng',
            'com.blackview.ai.vidgen',
            'com.blackview.ai.imagex',
            'com.blackview.ai.soundle',
            'com.blackview.ai.doki',
            'com.blackview.bvworkspace',
            'com.blackview.surfline',
            'com.blackview.dokegamecenter',

        ]

        packages_6300 = [
            'com.google.android.youtube',
            'com.google.android.apps.googleassistant',
            'com.google.android.apps.youtube.music',
            'com.google.android.keep',
            'com.android.music',
            'com.google.android.apps.maps',
            'com.mediatek.autodialer',
            'com.google.android.apps.photos',
            'com.wtk.gamemode',
            'com.android.manual',
            'com.google.android.apps.translate',
            'com.google.android.videos',
            'com.android.fmradio',
            'com.odm.sosapp',
            'com.hct.blackviewhome',
            'com.google.android.apps.tachyon',
            'com.opera.browser',
            'ru.yandex.yandexmaps',
            'ru.yandex.searchplugin',
            'com.yandex.browser',
            'com.blackview.easytrans',
            'com.odm.sosapp',
            'com.opera.preinstall',
            'com.opera.browser',
        ]

        if run_app != '\\':
            if run_app.startswith('BV6200'):

                print('Гаджет BV6200 найден, приступаю к работе...')
                for pkg in packages_6200:
                    print(f'{subprocess.call(["adb", "shell", "pm", "uninstall", "--user", "0", pkg])} Удален {pkg}\n')
                    time.sleep(0.5)

            if run_app.startswith('TE'):

                print('Гаджет BV6300 найден, приступаю к работе...')
                for pkg in packages_6300:
                    print(f'{subprocess.call(["adb", "shell", "pm", "uninstall", "--user", "0", pkg])} Удален {pkg}\n')
                    time.sleep(0.5)

            if run_app.startswith('BV4900'):

                print('Гаджет BV4900 найден, приступаю к работе...')
                for pkg in packages_6300:
                    print(f'{subprocess.call(["adb", "shell", "pm", "uninstall", "--user", "0", pkg])} Удален {pkg}\n')
                    time.sleep(0.5)

            commands = [
                'captive_portal_detection_enabled 0',
                'captive_portal_mode 0',
                'captive_portal_server https://192.168.206.2',
                'captive_portal_https_url https://192.168.206.2',
            ]

            for cmd in commands:
                subprocess.call(["adb", "shell", "settings", "put", "global", cmd])
                print(f'Команда: {cmd}')
                time.sleep(0.1)

            print('Проверка адреса сервера...')
            subprocess.call(["adb", "shell", "settings", "get", "global", "captive_portal_server"])

            print('Установка приложения...')
            device = subprocess.Popen(
                ['adb', 'devices'],
                stdout=subprocess.PIPE).communicate()[0].decode('cp866').split("\n")[1].split("\t")[0]
            subprocess.call(["adb", "-s", device, "install", "polymetal.apk"])
            print("Done, my Master!")
            time.sleep(30)
        else:
            print('Гаджет не найден')
            time.sleep(15)
